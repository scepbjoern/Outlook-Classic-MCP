"""Round trips through the registered MCP tools using fictional Outlook objects."""

import datetime as dt
from itertools import repeat
from types import SimpleNamespace
from unittest.mock import MagicMock, PropertyMock

import pytest

from outlook_mcp.utils.formatting import format_response
from tests.test_mailbox import Collection, make_profile, registered


STORE = {"store_id": "store-b"}
SERIES = {**STORE, "entry_id": "event-b", "scope": "series"}
UTC = dt.timezone.utc
REPLACE_TIME = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82280040"


def series():
    ns = make_profile()
    master = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    master.RecurrenceState = 1
    pattern = master.GetRecurrencePattern.return_value
    pattern.Exceptions = Collection()
    pattern.NoEndDate = True
    return ns, master, pattern


def occurrence(master, original, *, moved=None):
    event = MagicMock()
    event.EntryID, event.Parent, event.Class = master.EntryID, master.Parent, master.Class
    event.RecurrenceState = 2 if moved is None else 3
    event.Sensitivity, event.BusyStatus = 0, 2
    event.StartUTC = moved or original
    event.EndUTC = event.StartUTC + dt.timedelta(minutes=30)
    event.Subject = event.Body = "Fictional slot"
    event.Location = event.Organizer = event.Categories = ""
    event.ReminderMinutesBeforeStart, event.AllDayEvent = 10, False
    event.Recipients = [SimpleNamespace(Name="Guest", Address="guest@example.test", Type=1, MeetingResponseStatus=3)]
    event.MeetingStatus = 1
    if moved:
        event.PropertyAccessor.GetProperty.side_effect = {REPLACE_TIME: original}.__getitem__
    return event


def selected(row):
    return {field: row[field] for field in ("store_id", "entry_id", "scope", "original_start")}


def test_repeated_hour_and_moved_meeting_round_trip_to_different_items():
    ns, master, pattern = series()
    # Europe/Zurich's two local 02:30 occurrences on the fall-back date.
    early, late = (dt.datetime(2026, 10, 25, hour, 30, tzinfo=UTC) for hour in (0, 1))
    first = occurrence(master, early)
    moved = occurrence(master, late, moved=late + dt.timedelta(days=1))
    pattern.Exceptions.append(SimpleNamespace(Deleted=False, AppointmentItem=moved))
    pattern.GetOccurrence.return_value = first
    master.Parent.Items.Restrict.return_value = [first, moved]
    call = registered(ns)
    listing = call("list_events", **STORE, start="2026-10-25T02:30+02:00", end="2026-10-27T00:00+01:00")
    a, b = listing["items"]
    assert a["entry_id"] == b["entry_id"] == master.EntryID
    assert a["original_start"] == early.isoformat() and b["original_start"] == late.isoformat()
    assert not a["is_exception"] and b["is_exception"]
    assert dt.datetime.fromisoformat(b["start"]) == moved.StartUTC
    assert b["original_start"] in format_response(listing, "markdown")
    assert call("get_event", **selected(b))["attendees"][0]["address"] == "guest@example.test"
    call("update_event", **selected(a), subject="Only first")
    pattern.GetOccurrence.assert_called_with(early.astimezone().replace(tzinfo=UTC))
    call("delete_event", **selected(b))
    first.Save.assert_called_once_with()
    moved.Delete.assert_called_once_with()
    assert moved.Subject == "Fictional slot" and master.Subject == "b"
    master.Delete.assert_not_called()
    master.Save.assert_not_called()
    first.Send.assert_not_called()
    moved.Send.assert_not_called()


@pytest.mark.parametrize("args", [
    {}, {"scope": "occurrence"}, {"original_start": "2026-10-25T01:30Z"},
    {"scope": "series", "original_start": "2026-10-25T01:30Z"},
    {"scope": "occurrence", "original_start": "2026-10-25T02:30"},
])
@pytest.mark.parametrize("tool,fields", [
    ("update_event", {"subject": "Wrong"}), ("delete_event", {}),
    ("respond_event", {"response": "decline"}), ("set_category", {"categories": "Wrong"}),
])
def test_ambiguous_target_never_mutates(tool, fields, args):
    ns, master, _ = series()
    with pytest.raises(Exception, match="scope|original_start"):
        registered(ns)(tool, **STORE, entry_id=master.EntryID, **fields, **args)
    assert master.Subject == "b" and master.Categories == ""
    master.Delete.assert_not_called()
    master.Save.assert_not_called()
    master.Respond.assert_not_called()
    master.GetRecurrencePattern.assert_not_called()


@pytest.mark.parametrize("fault", ["deleted", "duplicate", "foreign_store", "wrong_id", "wrong_time",
                                   "master_returned", "outside_pattern", "unreadable", "unknown_state"])
def test_unprovable_occurrence_fails_before_assignment(fault):
    ns, master, pattern = series()
    slot = dt.datetime(2026, 11, 23, 10, tzinfo=UTC)
    event = occurrence(master, slot)
    pattern.GetOccurrence.return_value = event
    if fault == "deleted":
        entry = SimpleNamespace(Deleted=True, OriginalDate=slot.astimezone().replace(hour=0))
        pattern.Exceptions.append(entry)  # No AppointmentItem exists.
    elif fault == "duplicate":
        event = occurrence(master, slot, moved=slot + dt.timedelta(days=2))
        pattern.Exceptions.extend([SimpleNamespace(Deleted=False, AppointmentItem=event)] * 2)
    elif fault == "foreign_store":
        event.Parent = ns.Stores[0].GetDefaultFolder.return_value
    elif fault == "wrong_id":
        event.EntryID = "other-series"
    elif fault == "wrong_time":
        event.StartUTC += dt.timedelta(hours=1)
    elif fault == "master_returned":
        pattern.GetOccurrence.return_value = master
    elif fault == "outside_pattern":
        pattern.GetOccurrence.side_effect = RuntimeError("synthetic missing slot")
    elif fault == "unknown_state":
        master.RecurrenceState = False
    else:
        pattern.Exceptions = SimpleNamespace(Count=None)
    with pytest.raises(Exception, match="Item not found|scope"):
        registered(ns)("update_event", **STORE, entry_id=master.EntryID, scope="occurrence",
                       original_start=slot.isoformat(), subject="Wrong")
    assert event.Subject == "Fictional slot" and master.Subject == "b"
    event.Save.assert_not_called()
    master.Save.assert_not_called()


@pytest.mark.parametrize("exceptions", [Collection([SimpleNamespace(Deleted=True)]), SimpleNamespace(Count=None)])
def test_series_time_change_blocked_before_any_field_is_written(exceptions):
    ns, master, pattern = series()
    pattern.Exceptions = exceptions
    with pytest.raises(Exception, match="exceptions"):
        registered(ns)("update_event", **SERIES, subject="Wrong", end="2026-11-24T12:00")
    assert master.Subject == "b"
    master.Save.assert_not_called()


def test_explicit_series_writes_and_read_only_account():
    ns, master, _ = series()
    call = registered(ns)
    with pytest.raises(Exception, match="does not grant"):
        call("delete_event", **{**SERIES, "store_id": "store-a"})
    ns.GetItemFromID.assert_not_called()
    call("update_event", **SERIES, subject="Whole series")
    call("delete_event", **SERIES)
    assert master.Subject == "Whole series"
    master.Save.assert_called_once_with()
    master.Delete.assert_called_once_with()
    master.Send.assert_not_called()


def test_provider_reopens_exception_directly_and_preserves_offset_on_move():
    ns, master, _ = series()
    slot = dt.datetime(2026, 10, 25, 0, 30, tzinfo=UTC)
    item = occurrence(master, slot, moved=slot + dt.timedelta(days=1))
    item.EntryID = "provider-instance-id"
    ns.GetItemFromID.side_effect = lambda entry, store: item if (entry, store) == (item.EntryID, "store-b") else None
    call = registered(ns)
    args = {**STORE, "entry_id": item.EntryID, "scope": "occurrence", "original_start": slot.isoformat()}
    with pytest.raises(Exception, match="Item not found"):
        call("update_event", **{**args, "original_start": item.StartUTC.isoformat()}, subject="Wrong")
    with pytest.raises(Exception, match="master EntryID"):
        call("delete_event", **STORE, entry_id=item.EntryID, scope="series")
    item.Save.assert_not_called()
    item.Delete.assert_not_called()
    call("update_event", **args, start="2026-10-25T02:30:00+01:00", end="2026-10-25T03:00:00+01:00")
    assert item.StartUTC == dt.datetime(2026, 10, 25, 1, 30, tzinfo=UTC)
    assert item.EndUTC == dt.datetime(2026, 10, 25, 2, tzinfo=UTC)
    with pytest.raises(Exception, match="Category changes on individual series occurrences"):
        call("set_category", **args, categories="Synthetic category")
    assert item.Categories == master.Categories == ""
    item.Save.assert_called_once_with()  # Only the earlier time change.
    master.Save.assert_not_called()


def test_conflicting_provider_recurrence_flags_cannot_delete_as_single():
    ns, master, _ = series()
    master.RecurrenceState, master.IsRecurring = 0, True
    with pytest.raises(Exception, match="conflicting recurrence"):
        registered(ns)("delete_event", **STORE, entry_id=master.EntryID)
    master.Delete.assert_not_called()


def test_unreadable_original_has_no_actionable_list_id():
    ns, master, _ = series()
    slot = dt.datetime(2026, 11, 23, 8, tzinfo=UTC)
    item = occurrence(master, slot, moved=slot + dt.timedelta(hours=2))
    item.PropertyAccessor.GetProperty.side_effect = RuntimeError("unavailable replacement time")
    master.Parent.Items.Restrict.return_value = [item]
    row, = registered(ns)("list_events", **STORE, start="2026-11-22", end="2026-11-25")["items"]
    assert row["subject"] == "Fictional slot" and row["entry_id"] is None
    assert "scope" not in row and "original_start" not in row


@pytest.mark.parametrize("policy", ["busy_only", "exclude", "full"])
def test_scoped_private_exception_receipts_and_details(policy):
    ns, master, pattern = series()
    slot = dt.datetime(2026, 11, 23, 8, tzinfo=UTC)
    item = occurrence(master, slot, moved=slot + dt.timedelta(hours=4))
    item.Sensitivity = 2
    pattern.Exceptions.append(SimpleNamespace(Deleted=False, AppointmentItem=item))
    call = registered(ns, privacy={"default": policy})
    args = {**STORE, "entry_id": master.EntryID, "scope": "occurrence", "original_start": slot.isoformat()}
    if policy == "exclude":
        with pytest.raises(Exception, match="Item not found"):
            call("get_event", **args)
    else:
        result = call("get_event", **args)
        if policy == "busy_only":
            assert set(result) == {"subject", "busy_status", "start", "end"}
        else:
            assert selected(result) == args
    receipt = call("delete_event", **args)
    if policy != "full":
        assert receipt == {"status": "deleted"}
    item.Delete.assert_called_once_with()
    master.Delete.assert_not_called()


def test_indefinite_all_day_series_overlap_across_spring_dst():
    ns, master, pattern = series()
    # Local midnight to midnight in Europe/Zurich spans 23 hours.
    begin = dt.datetime(2026, 3, 28, 23, tzinfo=UTC)
    all_day = occurrence(master, begin)
    all_day.AllDayEvent, all_day.EndUTC = True, begin + dt.timedelta(hours=23)
    following = occurrence(master, all_day.EndUTC)
    items = master.Parent.Items
    def restricted(query):
        items.Sort.assert_called_once_with("[Start]")
        assert items.IncludeRecurrences is True and "[End] >=" in query and "[Start] <" in query
        return iter([all_day, following])
    items.Restrict.side_effect = restricted
    type(items).Count = PropertyMock(side_effect=AssertionError("undefined on expanded items"))
    result = registered(ns)("list_events", **STORE, start="2026-03-29T12:00+02:00", end="2026-03-30T00:00+02:00")
    row, = result["items"]
    assert row["all_day"] and row["original_start"] == begin.isoformat()
    pattern.GetOccurrence.assert_not_called()


def test_invalid_range_and_nonterminating_hidden_provider():
    ns, master, _ = series()
    call = registered(ns, privacy={"default": "exclude"})
    with pytest.raises(Exception, match="must follow"):
        call("list_events", **STORE, start="2026-10-26", end="2026-10-25")
    master.Parent.Items.Restrict.assert_not_called()
    hidden = SimpleNamespace(Parent=master.Parent, Sensitivity=2)
    master.Parent.Items.Restrict.return_value = repeat(hidden)
    with pytest.raises(Exception, match="narrower range"):
        call("list_events", **STORE, start="2026-10-25", end="2026-10-26")

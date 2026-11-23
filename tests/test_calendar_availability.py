"""Availability through the registered tools, with synthetic Outlook objects only."""

import datetime as dt
import json
from types import SimpleNamespace
from unittest.mock import PropertyMock

import pytest

from outlook_mcp.client import categories
from outlook_mcp.utils.formatting import format_response
from tests.test_calendar_categories import catalog
from tests.test_calendar_recurrence import SERIES, occurrence, selected, series
from tests.test_calendar_sensitivity import EXISTING, NEW_EVENT, STORE
from tests.test_mailbox import WINDOW, make_profile, registered


@pytest.mark.parametrize("code,label", list(enumerate(("Free", "Tentative", "Busy", "Out of office", "Working elsewhere"))))
@pytest.mark.parametrize("sensitivity", [0, 2], ids=["full-row", "reduced-row"])
def test_write_and_read_native_values(code, label, sensitivity):
    ns = make_profile()
    ev = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    call = registered(ns)
    saved = []
    ev.Save.side_effect = lambda: saved.append(ev.BusyStatus)
    call("create_event", **STORE, **NEW_EVENT, busy_status=code, sensitivity=sensitivity)
    for value in (code, (code + 1) % 5, code):
        call("update_event", **EXISTING, busy_status=value)
    assert saved == [code, code, (code + 1) % 5, code]
    listing = call("list_events", **STORE, **WINDOW)
    detail = call("get_event", **EXISTING)
    for payload in (listing, detail):
        assert f"{label} ({code})" in format_response(payload, "markdown")
        decoded = json.loads(format_response(payload, "json"))
        row = decoded["items"][0] if "items" in decoded else decoded
        assert row["busy_status"] == code
        if sensitivity == 2:
            assert set(row) == {"subject", "start", "end", "busy_status"}


@pytest.mark.parametrize("names,status,expected", [
    (None, None, 2), ([], None, 2), (["P.AL extra"], None, 2), (["P.AL"], None, 0),
    (["Other", "P.AL"], None, 0), (["P.AL", "Other"], 3, 3), (["Other"], 0, 0),
])
def test_creation_defaults_and_overrides(monkeypatch, names, status, expected):
    ns = make_profile()
    ns.Stores[1].Categories = catalog("P.AL", "P.AL extra", "Other")
    monkeypatch.setattr(categories, "_list_separator", lambda: ";")
    call = registered(ns)
    fields = {"busy_status": status} if status is not None else {}
    call("create_event", **NEW_EVENT, **STORE, categories=names, **fields)
    assert call("get_event", **EXISTING)["busy_status"] == expected


def test_category_and_sensitivity_edits_never_assign_busy_status(monkeypatch):
    ns = make_profile()
    ns.Stores[1].Categories = catalog("P.AL")
    monkeypatch.setattr(categories, "_list_separator", lambda: ";")
    ev = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    status = PropertyMock(return_value=4)
    type(ev).BusyStatus = status
    call = registered(ns)
    for names in (["P.AL"], []):
        call("update_event", **EXISTING, replace_categories=names, sensitivity=1)
        status.assert_not_called()
        assert call("get_event", **EXISTING)["busy_status"] == 4
        status.reset_mock()


@pytest.mark.parametrize("bad", [-1, 5, True, False, 2.0, "0", None])
@pytest.mark.parametrize("tool,args", [("create_event", NEW_EVENT), ("update_event", EXISTING)])
def test_invalid_explicit_status_rejected_before_outlook(tool, args, bad):
    ns = make_profile()
    with pytest.raises(Exception, match="busy_status"):
        registered(ns)(tool, **{**args, **STORE}, busy_status=bad)
    assert not ns.mock_calls


@pytest.mark.parametrize("state", ["missing", "error", "invalid", "boolean", "null"])
@pytest.mark.parametrize("sensitivity", [0, 2])
def test_unknown_status_on_reads(state, sensitivity):
    ns = make_profile()
    ev = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    ev.Sensitivity, ev.Categories = sensitivity, "P.AL"
    if state == "missing":
        del ev.BusyStatus
    elif state == "error":
        type(ev).BusyStatus = PropertyMock(side_effect=RuntimeError("Synthetic error"))
    else:
        ev.BusyStatus = {"invalid": 5, "boolean": False, "null": None}[state]
    call = registered(ns)
    listing = call("list_events", **STORE, **WINDOW)
    detail = call("get_event", **EXISTING)
    for row in (listing["items"][0], detail):
        assert row["busy_status"] is None
    for payload in (listing, detail):
        assert "Unknown" in format_response(payload, "markdown")
        assert '"busy_status": null' in format_response(payload, "json")


def test_regular_occurrence_and_moved_exception_keep_separate_status():
    ns, master, pattern = series()
    call = registered(ns)
    call("update_event", **SERIES, busy_status=3)
    assert call("get_event", **SERIES)["busy_status"] == 3
    slots = [dt.datetime(2026, 11, d, 8, tzinfo=dt.timezone.utc) for d in (22, 23, 24)]
    first, moved, last = (occurrence(master, slots[0]),
                          occurrence(master, slots[1], moved=slots[1] + dt.timedelta(hours=2)),
                          occurrence(master, slots[2]))
    first.BusyStatus, moved.BusyStatus, last.BusyStatus = 1, 4, 2
    master.Parent.Items.Restrict.return_value = [first, moved, last]
    pattern.Exceptions.append(SimpleNamespace(Deleted=False, AppointmentItem=moved))
    def rows():
        return call("list_events", **STORE, **WINDOW)["items"]
    a, b, c = rows()
    assert [r["busy_status"] for r in (a, b, c)] == [1, 4, 2]
    for event, row in ((first, a), (moved, b), (last, c)):
        pattern.GetOccurrence.return_value = event
        assert call("get_event", **selected(row))["busy_status"] == row["busy_status"]
    call("update_event", **selected(b), busy_status=0)
    assert call("get_event", **selected(b))["busy_status"] == 0
    assert [r["busy_status"] for r in rows()] == [1, 0, 2]
    pattern.GetOccurrence.return_value = first
    call("update_event", **selected(a), busy_status=4)
    assert [r["busy_status"] for r in rows()] == [4, 0, 2]
    assert call("get_event", **SERIES)["busy_status"] == 3
    last.Save.assert_not_called()
    for event in (master, first, moved):
        event.Save.assert_called_once_with()

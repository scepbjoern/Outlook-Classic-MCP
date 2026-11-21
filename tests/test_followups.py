"""Synthetic Outlook views. No Outlook process or user configuration is needed."""

import json
from datetime import date, datetime, timedelta, timezone
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from outlook_mcp.constants import OL_CLASS_MAIL, OL_CLASS_TASK, OL_FOLDER_TASKS, OL_FOLDER_TODO
from .test_mailbox import make_profile, registered


DAY = date(2026, 10, 25)
NO_DATE = datetime(4501, 1, 1)
REQUEST = {"mode": "on_date", "local_date": DAY.isoformat(), "store_id": "store-b"}


def followup(key, day=DAY, *, mail=False, store="store-b", **fields):
    props = dict(EntryID=key, Subject=key, Parent=SimpleNamespace(StoreID=store), Sensitivity=0)
    if mail:
        props.update(Class=OL_CLASS_MAIL, IsMarkedAsTask=True, TaskDueDate=day, TaskCompletedDate=NO_DATE,
                     PropertyAccessor=SimpleNamespace(GetProperty=Mock(return_value=2)))
    else:
        props.update(Class=OL_CLASS_TASK, Complete=False, DueDate=day)
    return SimpleNamespace(**(props | fields))


def set_views(store, todo, tasks):
    folders = {key: SimpleNamespace(StoreID=store.StoreID, Items=items) for key, items in
               ((OL_FOLDER_TODO, todo), (OL_FOLDER_TASKS, tasks)) if items is not None}
    store.GetDefaultFolder.side_effect = folders.get


def test_dates_status_dedup_and_pagination(monkeypatch):
    ns = make_profile()
    # Outlook supplies wall-clock fields. Converting either instant would change its day.
    early = datetime(2026, 10, 25, tzinfo=timezone(timedelta(hours=12)))
    late = datetime(2026, 10, 25, 23, 59, 59, tzinfo=timezone(timedelta(hours=-9)))
    mail = followup("A", early, mail=True)
    task = followup("B", late)
    closed_flag = followup("closed-flag", mail=True)
    closed_flag.PropertyAccessor.GetProperty.return_value = 1
    set_views(ns.Stores[1], [mail, task, closed_flag,
                             followup("completed", mail=True, TaskCompletedDate=DAY),
                             followup("unflagged", mail=True, IsMarkedAsTask=False),
                             followup("no-mail-date", NO_DATE, mail=True)],
              [task, mail, followup("done", Complete=True), followup("null-date", None),
               followup("sentinel", NO_DATE), followup("yesterday", DAY - timedelta(days=1)),
               followup("tomorrow", DAY + timedelta(days=1))])
    call = registered(ns)
    result = call("query_due", **REQUEST)
    assert result["coverage"] == {"to_do": "complete", "default_tasks": "complete",
                                  "calendar": "unsupported_no_native_due_date"}
    assert result["count"] == 2 and result["store_id"] == "store-b"
    assert result["items"] == [
        {"entry_id": key, "subject": key, "status": "outstanding", "due_date": DAY.isoformat(),
         "item_type": kind} for key, kind in (("A", "mail"), ("B", "task"))
    ]
    assert call("query_due", **REQUEST, limit=1)["next_offset"] == 1
    page = call("query_due", **REQUEST, limit=1, offset=1)
    assert page["items"] == [result["items"][1]] and page["next_offset"] is None

    class FixedDate(date):
        @classmethod
        def today(cls):
            return DAY

    monkeypatch.setattr("outlook_mcp.tools.followups.date", FixedDate)
    assert call("query_due", store_id="store-b")["items"] == result["items"]
    overdue = call("query_due", store_id="store-b", mode="overdue")
    assert [row["subject"] for row in overdue["items"]] == ["yesterday"]


@pytest.mark.parametrize("privacy", ["full", "busy_only", "exclude"])
def test_calendar_never_discloses_even_in_mixed_views(privacy):
    ns = make_profile()

    class CalendarRow:
        Class = 26

        def __getattr__(self, name):
            pytest.fail(f"Read appointment field {name}")

    hidden = followup("private-task", Sensitivity=2, Subject="hidden")
    unknown = followup("unknown-mail", mail=True, Sensitivity=None, Subject="hidden")
    set_views(ns.Stores[1], [CalendarRow(), unknown], [CalendarRow(), hidden])
    result = registered(ns, privacy={"default": privacy})("query_due", **REQUEST)
    assert result["count"] == 2 and "hidden" not in json.dumps(result)
    assert all(row["subject"] is None for row in result["items"])


def test_unreadable_sources_and_dates_are_not_reported_as_empty():
    ns = make_profile()
    store = ns.Stores[1]
    set_views(store, None, [])
    call = registered(ns)
    result = call("query_due", **REQUEST)
    assert result["count"] == 0 and result["coverage"]["to_do"] == "partial"
    assert result["coverage"]["default_tasks"] == "complete"

    def interrupted():
        yield followup("visible", mail=True)
        raise RuntimeError("private provider text")

    set_views(store, interrupted(), [followup("bad-date", "bad"), followup("bad-state", Complete=None)])
    result = call("query_due", **REQUEST)
    assert result["coverage"]["to_do"] == result["coverage"]["default_tasks"] == "partial"
    assert [row["subject"] for row in result["items"]] == ["visible"]
    assert "private provider text" not in json.dumps(result)


def test_explicit_read_authorization_and_store_proof():
    ns = make_profile()
    a, b = ns.Stores
    # Same EntryID is legal in two different stores.
    set_views(a, [followup("shared", store="store-a", Subject="A")], [])
    set_views(b, [followup("shared", Subject="B"), followup("foreign", store="store-a")], [])
    call = registered(ns, read=("b",), write=("a",))
    with pytest.raises(Exception, match="does not grant read"):
        call("query_due", **(REQUEST | {"store_id": "store-a"}))
    with pytest.raises(Exception, match="different Outlook stores"):
        call("query_due", **REQUEST, account="a@example.test")
    a.GetDefaultFolder.assert_not_called()
    result = call("query_due", **REQUEST)
    assert [row["subject"] for row in result["items"]] == ["B"]
    assert result["coverage"]["to_do"] == "partial"
    both = registered(ns)
    assert both("query_due", **(REQUEST | {"store_id": "store-a"}))["items"][0]["subject"] == "A"

    wrong = SimpleNamespace(StoreID="store-a", Items=[followup("must not be inspected")])
    b.GetDefaultFolder.side_effect = lambda kind: wrong
    result = call("query_due", **REQUEST)
    assert result["count"] == 0 and result["coverage"]["default_tasks"] == "partial"


@pytest.mark.parametrize("args", [
    {}, {"store_id": "store-a", "mode": "on_date"}, {"store_id": "store-a", "local_date": "2026-01-01"},
    {"store_id": "store-a", "mode": "on_date", "local_date": "20261025"},
    {"store_id": "store-a", "mode": "on_date", "local_date": "2026-02-30"},
])
def test_invalid_selection_or_date_before_store_lookup(args):
    ns = make_profile(("a",))
    with pytest.raises(Exception, match="local_date|Select the requested mailbox"):
        registered(ns)("query_due", **args)
    ns.Stores[0].GetDefaultFolder.assert_not_called()

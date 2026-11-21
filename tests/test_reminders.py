"""Current reminder states, date filtering and privacy at the MCP boundary."""

from datetime import datetime, timedelta
from types import SimpleNamespace

import pytest

from .test_mailbox import Collection, make_profile, registered
from .test_followups import DAY, REQUEST, set_views


def reminder(key, *, visible=True, delta=0, snooze=0, store="store-b", sensitivity=0, state=0):
    when = datetime.combine(DAY + timedelta(days=delta), datetime.min.time())
    event = SimpleNamespace(Class=26, Parent=SimpleNamespace(StoreID=store), RecurrenceState=state,
                            Sensitivity=sensitivity, Subject=key, Start=when - timedelta(days=9))
    return SimpleNamespace(Item=event, IsVisible=visible, NextReminderDate=when,
                           OriginalReminderDate=when - timedelta(minutes=snooze))


def test_current_reminders_survive_the_event_and_disappear_on_dismissal():
    ns = make_profile()
    active = reminder("visible", delta=-4)
    pending = reminder("pending", visible=False)
    snoozed = reminder("snoozed", visible=False, snooze=90)
    inaccessible_before = reminder("unknown-snooze", visible=False)
    inaccessible_before.OriginalReminderDate = None
    alerts = ns.Application.Reminders = Collection([active, pending, snoozed, inaccessible_before])
    call = registered(ns)
    def active_rows():
        return call("query_due", store_id="store-b", mode="active_reminders")

    assert [r["subject"] for r in active_rows()["items"]] == ["visible"]
    ns.Stores[1].GetDefaultFolder.assert_not_called()
    set_views(ns.Stores[1], [], [])
    dated = call("query_due", **REQUEST)
    assert {r["subject"]: r["status"] for r in dated["items"]} == {
        "snoozed": "snoozed", "pending": "scheduled", "unknown-snooze": "pending"}
    assert all(r["appointment_start"][:10] < DAY.isoformat() == r["due_date"] for r in dated["items"])
    assert all("entry_id" not in r for r in dated["items"])
    alerts.remove(active)  # The user dismisses one current reminder.
    assert active_rows()["count"] == 0
    assert call("query_due", **REQUEST)["items"] == dated["items"]


@pytest.mark.parametrize("privacy,expected", [("full", 1), ("busy_only", 0), ("exclude", 0)])
def test_reminders_prove_store_and_disclosure_before_reading_alert_or_title(privacy, expected):
    ns = make_profile()
    private = reminder("never disclose", sensitivity=2)
    unknown = reminder("unknown", sensitivity=None)

    class Foreign:
        Parent = SimpleNamespace(StoreID="store-a")

        def __getattr__(self, name):
            pytest.fail("Read foreign field " + name)

    ns.Application.Reminders = Collection([private, unknown, SimpleNamespace(Item=Foreign())])
    set_views(ns.Stores[1], [], [])
    result = registered(ns, privacy={"default": privacy})("query_due", **REQUEST)
    assert result["coverage"]["calendar"] == "complete"
    assert len(result["items"]) == expected
    assert all(r["subject"] is None for r in result["items"])


def test_unverifiable_store_and_unreadable_reminder_are_partial():
    ns = make_profile()
    set_views(ns.Stores[1], [], [])
    bad_store = reminder("secret", store=None)
    bad_date = reminder("no date")
    class VisibleWithoutTime:
        Item = bad_date.Item
        IsVisible = True

        @property
        def NextReminderDate(self):
            raise RuntimeError("unavailable")

    bad_date = VisibleWithoutTime()
    ns.Application.Reminders = Collection([bad_store, bad_date, reminder("ok")])
    call = registered(ns)
    result = call("query_due", **REQUEST)
    assert [r["subject"] for r in result["items"]] == ["ok"]
    assert result["coverage"]["calendar"] == "partial"
    result = call("query_due", mode="active_reminders", store_id="store-b")
    assert {r["subject"] for r in result["items"]} == {"ok", "no date"}


def test_master_cannot_contribute_a_stale_start_or_private_exception_title():
    ns = make_profile()
    master = reminder("master title", state=1)
    exceptions = Collection([SimpleNamespace(Deleted=False, AppointmentItem=SimpleNamespace(Sensitivity=2))])
    master.Item.GetRecurrencePattern = lambda: SimpleNamespace(Exceptions=exceptions)
    ns.Application.Reminders = Collection([master])
    args = {"mode": "active_reminders", "store_id": "store-b"}
    assert registered(ns)("query_due", **args)["count"] == 0
    row, = registered(ns, privacy={"default": "full"})("query_due", **args)["items"]
    assert row["appointment_start"] is None and row["subject"] is None

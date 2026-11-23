"""Native calendar properties through the registered tools, using fictional stores."""

import datetime as dt
import os
import time
from types import SimpleNamespace

import pytest

from tests.test_calendar_recurrence import SERIES, occurrence, series
from tests.test_calendar_sensitivity import EXISTING, STORE
from tests.test_mailbox import make_profile, registered


DATES = {"start": "2026-04-01", "end": "2026-04-04"}


@pytest.mark.parametrize("fields,expected", [({}, False), ({"all_day": False}, False), ({"all_day": True}, True)])
def test_midnight_span_and_unrelated_update(fields, expected):
    ns = make_profile()
    call = registered(ns)
    call("create_event", **STORE, subject="Synthetic", **DATES, **fields, reminder_minutes=None, busy_status=0)
    call("update_event", **EXISTING, subject="Renamed")
    detail = call("get_event", **EXISTING)
    assert detail["all_day"] is expected
    assert detail["reminder_set"] is False and detail["busy_status"] == 0
    for field in DATES:
        assert detail[field].startswith(DATES[field] + "T00:00:00")


@pytest.fixture
def zurich():
    # POSIX can select Zurich per process. On Windows exercise the real
    # machine zone when configured for the same Central European rules.
    previous = os.environ.get("TZ")
    try:
        if hasattr(time, "tzset"):
            os.environ["TZ"] = "Europe/Zurich"
            time.tzset()
        if [dt.datetime(2026, m, 1).astimezone().utcoffset() for m in (1, 7)] != [
            dt.timedelta(hours=1), dt.timedelta(hours=2)
        ]:
            pytest.skip("Needs a Central European Windows time zone or POSIX tzset")
        yield
    finally:
        if hasattr(time, "tzset"):
            if previous is None:
                os.environ.pop("TZ")
            else:
                os.environ["TZ"] = previous
            time.tzset()


@pytest.mark.parametrize("start,end,hours", [
    ("2026-03-29T00:00:00+01:00", "2026-03-30T00:00:00+02:00", 23),
    ("2026-10-25T00:00:00+02:00", "2026-10-26T00:00:00+01:00", 25),
    ("2026-03-28T00:00:00+01:00", "2026-03-31T00:00:00+02:00", 71),
])
def test_zurich_native_boundaries(zurich, start, end, hours):
    ns = make_profile()
    call = registered(ns)
    ev = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    call("create_event", **STORE, subject="Synthetic", start=start, end=end, all_day=True)
    assert ev.AllDayEvent is True
    detail = call("get_event", **EXISTING)
    assert (detail["start"], detail["end"]) == (start, end)
    assert dt.datetime.fromisoformat(end) - dt.datetime.fromisoformat(start) == dt.timedelta(hours=hours)


@pytest.mark.parametrize("start,end", [
    ("2026-04-01", "2026-04-01"), ("2026-04-02", "2026-04-01"),
    ("2026-04-01T12:00", "2026-04-02"), ("2026-04-01", "2026-04-02T00:00:01"),
    ("2026-02-30", "2026-04-01"),
])
def test_invalid_interval_rejected_before_mutation(start, end):
    ns = make_profile()
    items = ns.Stores[1].GetDefaultFolder.return_value.Items
    ev = items.Add.return_value
    call = registered(ns)
    with pytest.raises(Exception, match="midnight|exclusive|ISO-8601"):
        call("create_event", **STORE, subject="Wrong", start=start, end=end, all_day=True)
    items.Add.assert_not_called()
    with pytest.raises(Exception, match="midnight|exclusive|ISO-8601"):
        call("update_event", **EXISTING, subject="Wrong", start=start, end=end, all_day=True)
    assert ev.AllDayEvent is False and ev.Subject == "b"
    ev.Save.assert_not_called()


def test_conversions_and_partial_boundaries():
    ns = make_profile()
    call = registered(ns)
    ev = ns.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    ev.BusyStatus = 4
    with pytest.raises(Exception, match="midnight"):
        call("update_event", **EXISTING, all_day=True)
    call("update_event", **EXISTING, all_day=True, **DATES)
    call("update_event", **EXISTING, start="2026-04-02")
    with pytest.raises(Exception, match="exclusive"):
        call("update_event", **EXISTING, end="2026-04-02")
    assert ev.AllDayEvent is True and ev.End.day == 4
    call("update_event", **EXISTING, all_day=False, end="2026-04-02T09:30")
    detail = call("get_event", **EXISTING)
    assert detail["all_day"] is False and detail["busy_status"] == 4
    assert detail["start"].startswith("2026-04-02T00:00")
    assert detail["end"].startswith("2026-04-02T09:30")


def test_recurring_creation_conversion_and_exception_guard():
    ns, master, pattern = series()
    call = registered(ns)
    call("create_event", **STORE, subject="Synthetic series", all_day=True, **DATES,
         recurrence={"type": "weekly", "occurrences": 3})
    assert master.AllDayEvent is True and pattern.Occurrences == 3
    call("update_event", **SERIES, all_day=False)
    assert master.AllDayEvent is False
    pattern.Exceptions.append(SimpleNamespace(Deleted=True))
    with pytest.raises(Exception, match="exceptions"):
        call("update_event", **SERIES, all_day=True, subject="Wrong")
    assert master.Subject == "Synthetic series" and master.AllDayEvent is False


def test_occurrence_conversion_stays_in_selected_slot():
    ns, master, pattern = series()
    slot = dt.datetime(2026, 4, 1, 9, tzinfo=dt.timezone.utc)
    instance = occurrence(master, slot)
    pattern.GetOccurrence.return_value = instance
    registered(ns)("update_event", **EXISTING, scope="occurrence", original_start=slot.isoformat(),
                   all_day=True, **DATES)
    assert instance.AllDayEvent is True and master.AllDayEvent is False
    master.Save.assert_not_called()
    instance.Save.assert_called_once_with()

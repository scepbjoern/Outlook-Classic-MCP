"""Classify the actual Outlook row before reading any display fields."""

from __future__ import annotations

import datetime as dt
from typing import Any, NamedTuple

from outlook_mcp.client.folders import _safe_get
from outlook_mcp.config import PrivacyMode
from outlook_mcp.utils.formatting import to_iso


def native_sensitivity(event: Any) -> int | None:
    value = _safe_get(event, "Sensitivity")
    return value if type(value) is int and 0 <= value <= 3 else None


def _mode(sensitivity: int | None, policy: PrivacyMode) -> PrivacyMode:
    if sensitivity is None:
        return max(policy, PrivacyMode.BUSY_ONLY)
    return policy if sensitivity == 2 else PrivacyMode.FULL


def _series_mode(event: Any, policy: PrivacyMode, mode: PrivacyMode) -> PrivacyMode:
    unknown = _mode(None, policy)
    if mode == unknown:
        return mode
    state = _safe_get(event, "RecurrenceState")
    if type(state) is not int or state not in (0, 1, 2, 3):
        return unknown
    if state != 1:
        return mode

    # By-ID lookup can return a master for an exception's EntryID. A master
    # may be disclosed only as far as every live exception permits. Inspect
    # the finite Exceptions collection, never the expanded recurrence series.
    try:
        exceptions = event.GetRecurrencePattern().Exceptions
        count = exceptions.Count
        if type(count) is not int or count < 0:
            return unknown
        for index in range(1, count + 1):
            exception = exceptions.Item(index)
            deleted = exception.Deleted
            if deleted is True:
                continue
            if deleted is not False:
                return unknown
            mode = max(mode, _mode(native_sensitivity(exception.AppointmentItem), policy))
            if mode == unknown:
                break
    except Exception:
        return unknown
    return mode


class Disclosure(NamedTuple):
    mode: PrivacyMode
    sensitivity: int | None


def classify_event(event: Any, policy: PrivacyMode, *, requested: int | None = None) -> Disclosure:
    sensitivity = native_sensitivity(event)
    if requested is not None and requested != sensitivity:
        # A provider reporting a different value immediately after Save
        # cannot establish the sensitivity of the saved appointment.
        sensitivity = None
    return Disclosure(_series_mode(event, policy, _mode(sensitivity, policy)), sensitivity)


def availability(event: Any) -> dict[str, Any]:
    # Do not stringify malformed provider values into the allowed fields.
    start, end = _safe_get(event, "Start"), _safe_get(event, "End")
    busy = _safe_get(event, "BusyStatus")
    return {
        "subject": "Appointment",
        "start": to_iso(start) if isinstance(start, dt.datetime) else None,
        "end": to_iso(end) if isinstance(end, dt.datetime) else None,
        "busy_status": busy if type(busy) is int and 0 <= busy <= 4 else None,
    }

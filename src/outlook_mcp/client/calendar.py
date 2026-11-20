"""Calendar (appointment / meeting) COM operations."""

from __future__ import annotations

import datetime as dt
import sys
from typing import Any

from outlook_mcp.client.calendar_privacy import availability, classify_event, event_time
from outlook_mcp.client.calendar_target import exception_count, require_event_store, select_event, selection_fields
from outlook_mcp.client.categories import event_categories, require_category_scope
from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.config import PrivacyMode
from outlook_mcp.constants import (
    OL_APPOINTMENT_ITEM,
    OL_CLASS_APPOINTMENT,
    OL_FOLDER_CALENDAR,
    OL_MEETING,
    OL_MEETING_ACCEPTED,
    OL_MEETING_DECLINED,
    OL_MEETING_TENTATIVE,
    OL_RECURS_DAILY,
    OL_RECURS_MONTHLY,
    OL_RECURS_WEEKLY,
    OL_RECURS_YEARLY,
    OL_TO,
)
from outlook_mcp.errors import ITEM_NOT_FOUND, OutlookError, WriteOutcomeUnknown
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import from_iso, to_com, to_iso, truncate


_RECURRENCE_TYPE_MAP = {
    "daily": OL_RECURS_DAILY,
    "weekly": OL_RECURS_WEEKLY,
    "monthly": OL_RECURS_MONTHLY,
    "yearly": OL_RECURS_YEARLY,
}

_SENSITIVITY_LABELS = {0: "Normal", 1: "Personal", 2: "Private", 3: "Confidential"}


def _validate_sensitivity(value: int | None) -> None:
    if value is not None and (type(value) is not int or value not in _SENSITIVITY_LABELS):
        raise OutlookError("sensitivity must be an integer from 0 (Normal) to 3 (Confidential).")


def _updated_time(field: str, value: str) -> tuple[str, dt.datetime]:
    """An offset identifies an instant even when the local hour occurs twice."""
    try:
        parsed = dt.datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as exc:
        raise OutlookError("Use an ISO-8601 timestamp for appointment times.") from exc
    if parsed.tzinfo is not None:
        return field + "UTC", parsed.astimezone(dt.timezone.utc)
    return field, to_com(parsed)


def _event_summary(ev: Any, sensitivity: int | None) -> dict[str, Any]:
    recur_state = _safe_get(ev, "RecurrenceState", 0)
    return {
        "entry_id": _safe_get(ev, "EntryID"),
        "subject": _safe_get(ev, "Subject", ""),
        "start": event_time(ev, "Start"),
        "end": event_time(ev, "End"),
        "location": _safe_get(ev, "Location", ""),
        "organizer": _safe_get(ev, "Organizer", ""),
        "sensitivity": sensitivity,
        "sensitivity_label": _SENSITIVITY_LABELS.get(sensitivity, "Unknown"),
        "is_recurring": bool(recur_state),
        "all_day": bool(_safe_get(ev, "AllDayEvent", False)),
        "preview": truncate(_safe_get(ev, "Body", ""), 200),
    }


def _event_full(ev: Any, sensitivity: int | None) -> dict[str, Any]:
    attendees = []
    if _safe_get(ev, "Recipients"):
        for r in ev.Recipients:
            attendees.append(
                {
                    "name": _safe_get(r, "Name"),
                    "address": _safe_get(r, "Address"),
                    "type": _safe_get(r, "Type"),
                    "response": _safe_get(r, "MeetingResponseStatus"),
                }
            )
    return {
        **_event_summary(ev, sensitivity),
        "body": _safe_get(ev, "Body", ""),
        "attendees": attendees,
        "reminder_minutes": _safe_get(ev, "ReminderMinutesBeforeStart"),
        "categories": _safe_get(ev, "Categories", ""),
    }


def _apply_recurrence(appt: Any, spec: Recurrence) -> None:
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = _RECURRENCE_TYPE_MAP[spec.type]
    pattern.Interval = spec.interval
    if spec.occurrences is not None:
        pattern.Occurrences = spec.occurrences
    if spec.end_date is not None:
        pattern.PatternEndDate = to_com(from_iso(spec.end_date))


def _event_result(event: Any, mailbox: Mailbox, *, detail: bool = False) -> dict[str, Any] | None:
    mode, sensitivity = classify_event(event, mailbox.privacy)
    if mode == PrivacyMode.EXCLUDE:
        return None
    if mode == PrivacyMode.BUSY_ONLY:
        return availability(event)
    render = _event_full if detail else _event_summary
    return {**render(event, sensitivity), **selection_fields(event), "store_id": mailbox.store_id}


def _jet_date(day: dt.date) -> str:
    if sys.platform == "win32":
        import win32api

        # LOCALE_USER_DEFAULT, DATE_SHORTDATE. Jet uses the user's Windows
        # regional settings, which can differ from Python's process locale.
        clock = dt.datetime.combine(day, dt.time(), tzinfo=dt.timezone.utc)
        return win32api.GetDateFormat(0x0400, 1, clock, None).replace("'", "''")
    return day.strftime("%x")  # Synthetic non-Windows providers.


def list_events(
    namespace: Any,
    mailbox: Mailbox,
    *,
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    include_recurrences: bool = True,
) -> dict[str, Any]:
    try:
        start_dt = (dt.datetime.fromisoformat(start.replace("Z", "+00:00")) if start is not None
                    else dt.datetime.now()).astimezone()
        end_dt = (dt.datetime.fromisoformat(end.replace("Z", "+00:00")) if end is not None
                  else start_dt.replace(tzinfo=None) + dt.timedelta(days=14)).astimezone()
        # Jet compares local wall times with minute precision. Widen to
        # calendar days and refine by instant, also across the repeated hour.
        lower = start_dt.date()
        upper = end_dt.date() + dt.timedelta(days=1)
    except (ValueError, OSError, OverflowError) as exc:
        raise OutlookError("Use a valid ISO-8601 calendar range.") from exc
    if end_dt <= start_dt:
        raise OutlookError("The end of the calendar range must follow the start.")
    cal = mailbox.default_folder(OL_FOLDER_CALENDAR)
    items = cal.Items
    # Outlook requires the collection sorted ascending by [Start] BEFORE
    # IncludeRecurrences is set, or recurring series expand incorrectly.
    items.Sort("[Start]")
    items.IncludeRecurrences = include_recurrences

    restrict = (
        f"[End] >= '{_jet_date(lower)}' AND [Start] < '{_jet_date(upper)}'"
    )
    filtered = items.Restrict(restrict)

    results: list[dict[str, Any]] = []
    for scanned, ev in enumerate(filtered, 1):
        if scanned > 10_000:
            raise OutlookError("Cannot complete this calendar search. Use a narrower range.")
        require_event_store(ev, mailbox, namespace)
        row = _event_result(ev, mailbox)
        if row is None:
            continue
        begins, ends = row.get("start"), row.get("end")
        if begins is None or ends is None:
            continue
        begins, ends = dt.datetime.fromisoformat(begins), dt.datetime.fromisoformat(ends)
        if begins >= end_dt or (begins < start_dt and ends <= start_dt):
            continue
        results.append(row)
        if len(results) >= limit:
            break

    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "store_id": mailbox.store_id,
        "items": results,
    }


def _target(namespace: Any, mailbox: Mailbox, entry_id: str, scope: str | None,
            original_start: str | None, *, read: bool = False) -> Any:
    item = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_APPOINTMENT)
    return select_event(namespace, mailbox, item, entry_id=entry_id, scope=scope,
                        original_start=original_start, read=read)


def get_event(namespace: Any, mailbox: Mailbox, *, entry_id: str, scope: str | None = None,
              original_start: str | None = None) -> dict[str, Any]:
    try:
        event = _target(namespace, mailbox, entry_id, scope, original_start, read=True)
    except OutlookError as exc:
        if scope is not None or original_start is not None:
            raise OutlookError(ITEM_NOT_FOUND) from exc
        raise
    result = _event_result(event, mailbox, detail=True)
    if result is None:
        raise OutlookError(ITEM_NOT_FOUND)
    return result


def create_event(
    namespace: Any,
    mailbox: Mailbox,
    *,
    subject: str,
    start: str,
    end: str,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    is_online_meeting: bool = False,
    reminder_minutes: int | None = 15,
    recurrence: Recurrence | None = None,
    sensitivity: int | None = None,
    categories: list[str] | None = None,
) -> dict[str, Any]:
    _validate_sensitivity(sensitivity)
    mailbox.require_smtp_recipients(attendees or [])
    start_time = to_com(from_iso(start))
    end_time = to_com(from_iso(end))
    assignment = event_categories(namespace, mailbox, replace=categories)
    appt = mailbox.default_folder(OL_FOLDER_CALENDAR).Items.Add(OL_APPOINTMENT_ITEM)
    mailbox.require_store(appt.Parent)
    if assignment is not None:
        appt.Categories = assignment
    appt.Subject = subject
    # to_com, not the bare datetime: pywin32 would read a naive value
    # as local and hand Outlook its UTC clock, shifting the event one
    # offset early.
    appt.Start = start_time
    appt.End = end_time
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    if attendees:
        appt.MeetingStatus = OL_MEETING
        mailbox.bind_sender(appt)
        for addr in attendees:
            rec = appt.Recipients.Add(addr)
            rec.Type = OL_TO
        if not appt.Recipients.ResolveAll():
            raise OutlookError("Outlook could not resolve all meeting attendees.")
    if recurrence is not None:
        _apply_recurrence(appt, recurrence)
    if sensitivity is not None:
        appt.Sensitivity = sensitivity
    appt.Save()
    try:
        result = {"status": "created", "local_outcome": "saved"}
        mode = classify_event(appt, mailbox.privacy, requested=sensitivity).mode
        if mode == PrivacyMode.FULL:
            result.update(entry_id=appt.EntryID, store_id=mailbox.store_id, subject=appt.Subject,
                          start=to_iso(appt.Start), end=to_iso(appt.End))
        elif mode == PrivacyMode.BUSY_ONLY:
            result.update(availability(appt))
        if attendees:
            mailbox.bind_sender(appt)
            appt.Send()
            result["local_outcome"] = "saved_and_submitted"
    except Exception as exc:
        raise WriteOutcomeUnknown(
            "Outlook saved the appointment locally. A later step failed. "
            "Check the calendar and any invitation in Outbox or Sent Items."
        ) from exc
    return result


def update_event(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    scope: str | None = None,
    original_start: str | None = None,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    sensitivity: int | None = None,
    add_categories: list[str] | None = None,
    remove_categories: list[str] | None = None,
    replace_categories: list[str] | None = None,
) -> dict[str, Any]:
    _validate_sensitivity(sensitivity)
    ev = _target(namespace, mailbox, entry_id, scope, original_start)
    if any(names is not None for names in (add_categories, remove_categories, replace_categories)):
        require_category_scope(scope)
    times = [_updated_time(field, value) for field, value in (("Start", start), ("End", end))
             if value is not None]
    if scope == "series" and (start is not None or end is not None):
        try:
            count = exception_count(ev.GetRecurrencePattern().Exceptions)
        except Exception as exc:
            raise OutlookError("Cannot verify exceptions before changing series times.") from exc
        if count:
            raise OutlookError("Changing series times could discard exceptions. Select one occurrence.")
    assignment = event_categories(namespace, mailbox, event=ev, replace=replace_categories,
                                  add=add_categories, remove=remove_categories)
    if assignment is not None:
        ev.Categories = assignment
    if subject is not None:
        ev.Subject = subject
    for field, value in times:
        setattr(ev, field, value)
    if location is not None:
        ev.Location = location
    if body is not None:
        ev.Body = body
    if sensitivity is not None:
        ev.Sensitivity = sensitivity
    ev.Save()
    if classify_event(ev, mailbox.privacy, requested=sensitivity).mode != PrivacyMode.FULL:
        return {"status": "updated"}
    return {"status": "updated", "entry_id": entry_id, "store_id": mailbox.store_id}


def delete_event(namespace: Any, mailbox: Mailbox, *, entry_id: str, scope: str | None = None,
                 original_start: str | None = None) -> dict[str, Any]:
    ev = _target(namespace, mailbox, entry_id, scope, original_start)
    result = {"status": "deleted"}
    if classify_event(ev, mailbox.privacy).mode == PrivacyMode.FULL:
        result.update(subject=_safe_get(ev, "Subject", ""), entry_id=entry_id, store_id=mailbox.store_id)
    ev.Delete()
    return result


def respond_event(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    scope: str | None = None,
    original_start: str | None = None,
    response: str,
    send_response: bool = True,
) -> dict[str, Any]:
    code = {
        "accept": OL_MEETING_ACCEPTED,
        "tentative": OL_MEETING_TENTATIVE,
        "decline": OL_MEETING_DECLINED,
    }.get(response.lower())
    if code is None:
        raise OutlookError("response must be one of: 'accept', 'tentative', 'decline'.")
    ev = _target(namespace, mailbox, entry_id, scope, original_start)
    result = {"status": "responded"}
    if classify_event(ev, mailbox.privacy).mode == PrivacyMode.FULL:
        result.update(response=response, store_id=mailbox.store_id)
    resp = ev.Respond(code, True, False)
    if send_response:
        if resp is None:
            raise WriteOutcomeUnknown("Outlook updated the meeting without returning a response to send.")
        try:
            mailbox.bind_sender(resp)
            resp.Send()
        except Exception as exc:
            raise WriteOutcomeUnknown("Outlook processed the meeting response. Its submission is uncertain.") from exc
        result["local_outcome"] = "submitted"
    return result

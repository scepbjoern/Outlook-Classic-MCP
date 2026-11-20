"""Calendar (appointment / meeting) COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.client.mailbox import Mailbox
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
from outlook_mcp.errors import OutlookError, is_disconnect_error
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import from_iso, to_com, to_iso, truncate


_RECURRENCE_TYPE_MAP = {
    "daily": OL_RECURS_DAILY,
    "weekly": OL_RECURS_WEEKLY,
    "monthly": OL_RECURS_MONTHLY,
    "yearly": OL_RECURS_YEARLY,
}

_SENSITIVITY_LABELS = {0: "Normal", 1: "Personal", 2: "Private", 3: "Confidential"}

# PidTagStoreEntryId identifies the message store containing the item.
_PR_STORE_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFB0102"


def _validate_sensitivity(value: int | None) -> None:
    if value is not None and (type(value) is not int or value not in _SENSITIVITY_LABELS):
        raise OutlookError("sensitivity must be an integer from 0 (Normal) to 3 (Confidential).")


def _require_listed_event_store(ev: Any, mailbox: Mailbox, namespace: Any) -> None:
    parent_id = getattr(ev.Parent, "StoreID", None)
    if isinstance(parent_id, str) and parent_id:
        if not mailbox.matches_store_id(parent_id):
            raise OutlookError("Outlook returned a calendar item outside the selected store.")
        return

    # Some calendar rows have no folder StoreID on Parent. Verify that same
    # row's own MAPI identity before accessing its subject or other content.
    try:
        accessor = ev.PropertyAccessor
        item_store = accessor.BinaryToString(accessor.GetProperty(_PR_STORE_ENTRYID))
        if isinstance(item_store, str) and item_store:
            if mailbox.matches_store_id(item_store):
                return
            if namespace.CompareEntryIDs(item_store, mailbox.store_id) is True:
                return
    except Exception as exc:
        if is_disconnect_error(exc):
            raise
        raise OutlookError("Cannot verify the calendar item's message store.") from exc
    raise OutlookError("Cannot verify the calendar item belongs to the selected store.")


def _event_summary(ev: Any) -> dict[str, Any]:
    recur_state = _safe_get(ev, "RecurrenceState", 0)
    sensitivity = _safe_get(ev, "Sensitivity")
    return {
        "entry_id": _safe_get(ev, "EntryID"),
        "subject": _safe_get(ev, "Subject", ""),
        "start": to_iso(_safe_get(ev, "Start")),
        "end": to_iso(_safe_get(ev, "End")),
        "location": _safe_get(ev, "Location", ""),
        "organizer": _safe_get(ev, "Organizer", ""),
        "sensitivity": sensitivity,
        "sensitivity_label": _SENSITIVITY_LABELS.get(sensitivity, "Unknown"),
        "is_recurring": bool(recur_state),
        "all_day": bool(_safe_get(ev, "AllDayEvent", False)),
        "preview": truncate(_safe_get(ev, "Body", ""), 200),
    }


def _event_full(ev: Any) -> dict[str, Any]:
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
        **_event_summary(ev),
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


def list_events(
    namespace: Any,
    mailbox: Mailbox,
    *,
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    include_recurrences: bool = True,
) -> dict[str, Any]:
    cal = mailbox.default_folder(OL_FOLDER_CALENDAR)
    items = cal.Items
    # Outlook requires the collection sorted ascending by [Start] BEFORE
    # IncludeRecurrences is set, or recurring series expand incorrectly.
    items.Sort("[Start]")
    if include_recurrences:
        items.IncludeRecurrences = True

    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))

    # Jet filter dates must be 12-hour + AM/PM; %H with %p would emit
    # e.g. "14:30 PM", which Outlook misparses for afternoon times.
    restrict = (
        f"[Start] >= '{start_dt.strftime('%m/%d/%Y %I:%M %p')}' AND "
        f"[Start] <= '{end_dt.strftime('%m/%d/%Y %I:%M %p')}'"
    )
    filtered = items.Restrict(restrict)

    results: list[dict[str, Any]] = []
    for ev in filtered:
        _require_listed_event_store(ev, mailbox, namespace)
        results.append({**_event_summary(ev), "store_id": mailbox.store_id})
        if len(results) >= limit:
            break

    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "store_id": mailbox.store_id,
        "items": results,
    }


def get_event(namespace: Any, mailbox: Mailbox, *, entry_id: str) -> dict[str, Any]:
    event = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_APPOINTMENT)
    return {**_event_full(event), "store_id": mailbox.store_id}


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
) -> dict[str, Any]:
    _validate_sensitivity(sensitivity)
    mailbox.require_smtp_recipients(attendees or [])
    start_time = to_com(from_iso(start))
    end_time = to_com(from_iso(end))
    appt = mailbox.default_folder(OL_FOLDER_CALENDAR).Items.Add(OL_APPOINTMENT_ITEM)
    mailbox.require_store(appt.Parent)
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
    result = {
        "status": "created",
        "entry_id": appt.EntryID,
        "store_id": mailbox.store_id,
        "subject": appt.Subject,
        "start": to_iso(appt.Start),
        "end": to_iso(appt.End),
    }
    if attendees:
        mailbox.bind_sender(appt)
        appt.Send()
    return result


def update_event(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    sensitivity: int | None = None,
) -> dict[str, Any]:
    _validate_sensitivity(sensitivity)
    ev = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_APPOINTMENT)
    start_time = to_com(from_iso(start))
    end_time = to_com(from_iso(end))
    if subject is not None:
        ev.Subject = subject
    if start is not None:
        ev.Start = start_time
    if end is not None:
        ev.End = end_time
    if location is not None:
        ev.Location = location
    if body is not None:
        ev.Body = body
    if sensitivity is not None:
        ev.Sensitivity = sensitivity
    ev.Save()
    return {"status": "updated", "entry_id": entry_id, "store_id": mailbox.store_id}


def delete_event(namespace: Any, mailbox: Mailbox, *, entry_id: str) -> dict[str, Any]:
    ev = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_APPOINTMENT)
    subject = _safe_get(ev, "Subject", "")
    ev.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id, "store_id": mailbox.store_id}


def respond_event(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
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
    ev = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_APPOINTMENT)
    resp = ev.Respond(code, True, False)
    if send_response:
        if resp is None:
            raise OutlookError("Outlook updated the meeting without returning a response to send.")
        mailbox.bind_sender(resp)
        resp.Send()
    return {"status": "responded", "response": response, "store_id": mailbox.store_id}

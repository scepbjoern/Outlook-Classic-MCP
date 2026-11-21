"""Project Outlook's current reminder collection into one authorized mailbox."""

from datetime import date, datetime

from outlook_mcp.client.calendar_privacy import classify_event, event_time
from outlook_mcp.client.calendar_target import MASTER, require_event_store
from outlook_mcp.config import PrivacyMode
from outlook_mcp.constants import OL_CLASS_APPOINTMENT
from outlook_mcp.errors import OutlookError, is_disconnect_error
from outlook_mcp.utils.formatting import to_iso


def _clock(value):
    if not isinstance(value, datetime) or value.date() == date(4501, 1, 1):
        return None
    return value.replace(tzinfo=None)


def _reminder_row(reminder, mailbox, namespace, mode, target):
    event = reminder.Item
    parent_id = getattr(event.Parent, "StoreID", None)
    if isinstance(parent_id, str) and parent_id:
        if not mailbox.matches_store_id(parent_id):
            return None
    else:
        require_event_store(event, mailbox, namespace)
    if event.Class != OL_CLASS_APPOINTMENT:
        return None
    disclosure = classify_event(event, mailbox.privacy)
    if disclosure.mode != PrivacyMode.FULL:
        # Busy-only permits availability fields, which cannot disclose alert state.
        return None
    visible = reminder.IsVisible
    if type(visible) is not bool:
        raise ValueError("Unreadable reminder visibility")
    if mode == "active_reminders" and not visible:
        return None
    try:
        next_time = _clock(reminder.NextReminderDate)
    except Exception as exc:
        if mode != "active_reminders" or is_disconnect_error(exc):
            raise
        next_time = None
    if next_time is None and mode != "active_reminders":
        raise ValueError("Unreadable reminder time")
    day = next_time.date() if next_time else None
    if mode != "active_reminders" and not (day < target if mode == "overdue" else day == target):
        return None
    status = "active"
    if not visible:
        original = _clock(reminder.OriginalReminderDate)
        if original is None:
            status = "pending"
        else:
            status = "snoozed" if next_time > original else "scheduled"
    # Reminder.Item can be a series master. Its Start and Subject may belong
    # to a different occurrence. Never present those as instance details.
    master = event.RecurrenceState == MASTER
    return {
        "item_type": "calendar", "subject": event.Subject if not master and disclosure.sensitivity != 2 else None,
        "status": status, "due_date": day.isoformat() if day else None,
        "reminder_at": to_iso(next_time), "appointment_start": None if master else event_time(event, "Start"),
    }


def appointment_reminders(namespace, mailbox, mode, target):
    rows = []
    coverage = "complete"
    try:
        reminders = namespace.Application.Reminders
        count = reminders.Count
        if type(count) is not int or count < 0:
            raise ValueError("Unreadable reminder collection")
        for index in range(1, count + 1):
            try:
                row = _reminder_row(reminders.Item(index), mailbox, namespace, mode, target)
                if row is not None:
                    rows.append(row)
            except Exception as exc:
                if is_disconnect_error(exc):
                    raise
                coverage = "partial"
        if reminders.Count != count:
            coverage = "partial"
    except Exception as exc:
        if is_disconnect_error(exc):
            raise OutlookError("Outlook disconnected while reading reminders.") from exc
        coverage = "partial"
    return rows, coverage

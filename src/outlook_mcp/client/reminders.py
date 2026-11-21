"""Project Outlook's current reminder collection into one authorized mailbox."""

from datetime import date, datetime

from outlook_mcp.client.calendar_privacy import classify_event, event_time, native_sensitivity
from outlook_mcp.client.calendar_target import MASTER, mapi_store_id
from outlook_mcp.config import PrivacyMode, normalize_identifier
from outlook_mcp.constants import OL_CLASS_APPOINTMENT, OL_CLASS_MAIL, OL_CLASS_TASK
from outlook_mcp.errors import OutlookError, is_disconnect_error
from outlook_mcp.utils.formatting import to_iso


def _clock(value):
    if not isinstance(value, datetime) or value.date() == date(4501, 1, 1):
        return None
    return value.replace(tzinfo=None)


def reminder_store_id(item, namespace):
    parent_id = getattr(item.Parent, "StoreID", None)
    if isinstance(parent_id, str) and parent_id:
        return parent_id
    # Some reminders expose only a MAPI store identifier, which can be
    # encoded differently from Store.StoreID. Resolve it before reading Class
    # or any display properties, even when the owner is outside the request.
    raw = mapi_store_id(item)
    if isinstance(raw, str) and raw:
        stores = namespace.Stores
        for index in range(1, stores.Count + 1):
            known = stores.Item(index).StoreID
            if isinstance(known, str) and known and (
                normalize_identifier(raw) == normalize_identifier(known)
                or namespace.CompareEntryIDs(raw, known) is True
            ):
                return known
    raise OutlookError("Cannot identify the reminder's store.")


ITEM_TYPES = {OL_CLASS_APPOINTMENT: "calendar", OL_CLASS_MAIL: "mail", OL_CLASS_TASK: "task"}


def reminder_row(reminder, item, mailbox, item_type, mode, target=None):
    if item_type == "calendar":
        disclosure = classify_event(item, mailbox.privacy)
        if disclosure.mode != PrivacyMode.FULL:
            # Availability alone cannot disclose alert state.
            return None
        sensitivity = disclosure.sensitivity
    else:
        sensitivity = native_sensitivity(item)
    visible = reminder.IsVisible
    if type(visible) is not bool:
        raise ValueError("Unreadable reminder visibility")
    if mode == "active_reminders" and not visible:
        return None
    try:
        next_time = _clock(reminder.NextReminderDate)
    except Exception as exc:
        if mode not in ("active_reminders", "all") or is_disconnect_error(exc):
            raise
        next_time = None
    if next_time is None and mode not in ("active_reminders", "all"):
        raise ValueError("Unreadable reminder time")
    day = next_time.date() if next_time else None
    if mode not in ("active_reminders", "all") and not (day < target if mode == "overdue" else day == target):
        return None
    status = "active"
    if not visible:
        original = _clock(reminder.OriginalReminderDate)
        if original is None or next_time is None:
            status = "pending"
        else:
            status = "snoozed" if next_time > original else "scheduled"
    # Reminder.Item can be a series master. Its Start and Subject may belong
    # to a different occurrence. Never present those as instance details.
    master = item_type == "calendar" and item.RecurrenceState == MASTER
    return {
        "item_type": item_type, "subject": item.Subject if not master and sensitivity in (0, 1, 3) else None,
        "status": status, "due_date": day.isoformat() if day else None,
        "reminder_at": to_iso(next_time), "appointment_start": event_time(item, "Start") if item_type == "calendar" and not master else None,
    }


def scan_reminders(namespace, project, *, collection=None):
    rows = []
    coverage = "complete"
    try:
        reminders = namespace.Application.Reminders if collection is None else collection
        count = reminders.Count
        if type(count) is not int or count < 0:
            raise ValueError("Unreadable reminder collection")
        for index in range(1, count + 1):
            try:
                row = project(reminders.Item(index))
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


def appointment_reminders(namespace, mailbox, mode, target):
    def project(reminder):
        item = reminder.Item
        if mailbox.matches_store_id(reminder_store_id(item, namespace)) and item.Class == OL_CLASS_APPOINTMENT:
            return reminder_row(reminder, item, mailbox, "calendar", mode, target)
        return None

    return scan_reminders(namespace, project)

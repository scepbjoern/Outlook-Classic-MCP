"""Read outstanding follow-ups from the selected store's local Outlook views."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from outlook_mcp.client.calendar_privacy import native_sensitivity
from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.constants import OL_CLASS_MAIL, OL_CLASS_TASK, OL_FOLDER_TASKS, OL_FOLDER_TODO
from outlook_mcp.errors import OutlookError, is_disconnect_error


def _local_day(value: Any) -> date | None:
    # Outlook COM dates carry local wall-clock fields, even with a UTC tzinfo.
    if isinstance(value, datetime):
        value = value.date()
    if value is None or value == date(4501, 1, 1):
        return None
    if not isinstance(value, date):
        raise ValueError("Unreadable Outlook date")
    return value


def _outstanding_due(item: Any, item_class: int) -> date | None:
    if item_class == OL_CLASS_TASK:
        complete = item.Complete
        if complete is True:
            return None
        if complete is not False:
            raise ValueError("Unknown completion state")
        return _local_day(item.DueDate)
    marked = item.IsMarkedAsTask
    if marked is False:
        return None
    if marked is not True:
        raise ValueError("Unknown flag state")
    flag = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10900003")
    if type(flag) is not int or flag not in (1, 2):
        raise ValueError("Unknown native flag status")
    if flag == 1 or _local_day(item.TaskCompletedDate) is not None:
        return None
    return _local_day(item.TaskDueDate)


def _matching_row(item, mailbox, target, overdue):
    item_class = item.Class
    if type(item_class) is not int:
        raise ValueError("Unknown item class")
    if item_class not in (OL_CLASS_TASK, OL_CLASS_MAIL):
        return None
    mailbox.require_store(item.Parent)
    due = _outstanding_due(item, item_class)
    if due is None or not (due < target if overdue else due == target):
        return None
    entry_id = item.EntryID
    if not isinstance(entry_id, str) or not entry_id:
        raise ValueError("Missing item identifier")
    # Unknown and Private sensitivity never contribute a title to this query.
    subject = item.Subject if native_sensitivity(item) in (0, 1, 3) else None
    return {
        "entry_id": entry_id, "item_type": "task" if item_class == OL_CLASS_TASK else "mail",
        "subject": subject, "due_date": due.isoformat(), "status": "outstanding",
    }


def query_due(namespace: Any, mailbox: Mailbox, *, mode: str, target: date,
              offset: int, limit: int) -> dict[str, Any]:
    # The To-Do folder can also surface tasks in the ordinary Tasks folder.
    rows = {}
    coverage = {"calendar": "unsupported_no_native_due_date"}
    for source, folder_type in (("to_do", OL_FOLDER_TODO), ("default_tasks", OL_FOLDER_TASKS)):
        coverage[source] = "complete"
        try:
            folder = mailbox.default_folder(folder_type)
            for item in folder.Items:
                try:
                    row = _matching_row(item, mailbox, target, mode == "overdue")
                    if row is not None:
                        rows[row["entry_id"].lower()] = row
                except Exception as exc:
                    if is_disconnect_error(exc):
                        raise
                    coverage[source] = "partial"
        except Exception as exc:
            if is_disconnect_error(exc):
                raise OutlookError("Outlook disconnected while reading due items.") from exc
            coverage[source] = "partial"
    ordered = sorted(rows.values(), key=lambda row: (row["due_date"], row["entry_id"]))
    page = ordered[offset:offset + limit]
    end = offset + len(page)
    return {
        "store_id": mailbox.store_id, "mode": mode, "local_date": target.isoformat(),
        "coverage": coverage, "count": len(page), "items": page,
        "next_offset": end if end < len(ordered) else None,
    }

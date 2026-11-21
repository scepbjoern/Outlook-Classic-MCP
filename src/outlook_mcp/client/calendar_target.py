"""Resolve a calendar selection inside the already authorized Outlook store."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get
from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.constants import OL_CLASS_APPOINTMENT
from outlook_mcp.errors import ITEM_NOT_FOUND, OutlookError, is_disconnect_error
from outlook_mcp.utils.formatting import to_com


SINGLE, MASTER, OCCURRENCE, EXCEPTION = range(4)  # OlRecurrenceState
_REPLACE_TIME = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/82280040"
_STORE_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x0FFB0102"


def mapi_store_id(item: Any) -> str:
    accessor = item.PropertyAccessor
    return accessor.BinaryToString(accessor.GetProperty(_STORE_ENTRYID))


def require_event_store(event: Any, mailbox: Mailbox, namespace: Any) -> None:
    parent = event.Parent
    parent_id = getattr(parent, "StoreID", None)
    if isinstance(parent_id, str) and parent_id:
        mailbox.require_store(parent)
        return
    # Some expanded items have only a MAPI store identity.
    store = mapi_store_id(event)
    if isinstance(store, str) and store:
        if mailbox.matches_store_id(store) or namespace.CompareEntryIDs(store, mailbox.store_id) is True:
            return
    raise OutlookError("Cannot verify the calendar item belongs to the selected store.")


def _original_start(event: Any, state: int) -> dt.datetime:
    value = (event.PropertyAccessor.GetProperty(_REPLACE_TIME)
             if state == EXCEPTION else event.StartUTC)
    if not isinstance(value, dt.datetime):
        raise OutlookError("Outlook could not verify the original occurrence time.")
    # Both COM properties contain UTC clock fields, regardless of pywin32's tzinfo.
    return value.replace(tzinfo=dt.timezone.utc)


def selection_fields(event: Any) -> dict[str, Any]:
    state = _safe_get(event, "RecurrenceState")
    if type(state) is int and state == MASTER:
        return {"scope": "series"}
    if type(state) is int and state in (OCCURRENCE, EXCEPTION):
        try:
            original = _original_start(event, state)
        except Exception as exc:
            if is_disconnect_error(exc):
                raise
            # A displayed row without a verified slot must not open the master.
            return {"entry_id": None}
        return {"scope": "occurrence", "original_start": original.isoformat(),
                "is_exception": state == EXCEPTION}
    return {}


def exception_count(exceptions: Any) -> int:
    count = exceptions.Count
    if type(count) is not int or count < 0:
        raise OutlookError("Outlook could not verify the series exceptions.")
    return count


def _resolve_occurrence(namespace: Any, mailbox: Mailbox, master: Any, original: dt.datetime) -> Any:
    pattern = master.GetRecurrencePattern()
    local = original.astimezone().replace(tzinfo=None)
    exceptions = pattern.Exceptions
    match = None
    for index in range(1, exception_count(exceptions) + 1):
        exception = exceptions.Item(index)
        deleted = exception.Deleted
        if deleted is True:
            # Outlook drops the time component after deletion.
            if exception.OriginalDate.date() == local.date():
                raise OutlookError(ITEM_NOT_FOUND)
            continue
        if deleted is not False:
            raise OutlookError(ITEM_NOT_FOUND)
        item = exception.AppointmentItem
        require_event_store(item, mailbox, namespace)
        if _original_start(item, EXCEPTION) == original:
            if match is not None:
                raise OutlookError(ITEM_NOT_FOUND)
            match = item

    if match is None:
        item = pattern.GetOccurrence(to_com(local))
        require_event_store(item, mailbox, namespace)
        expected = OCCURRENCE
    else:
        item, expected = match, EXCEPTION
    state = item.RecurrenceState
    if (item.Class != OL_CLASS_APPOINTMENT or type(state) is not int or state != expected
            or _original_start(item, state) != original):
        raise OutlookError(ITEM_NOT_FOUND)
    # This series' pattern is the source of the item. Reject a different
    # EntryID as well: a provider with unverified identity fails closed.
    if item.EntryID != master.EntryID:
        raise OutlookError(ITEM_NOT_FOUND)
    return item


def select_event(namespace: Any, mailbox: Mailbox, item: Any, *, entry_id: str, scope: str | None,
                 original_start: str | None, read: bool = False) -> Any:
    if not entry_id.strip() or item.EntryID != entry_id:
        raise OutlookError(ITEM_NOT_FOUND)
    if scope is None and original_start is None and read:
        return item  # Legacy, privacy-filtered detail of the object Outlook opened.
    state = _safe_get(item, "RecurrenceState")
    recurring = _safe_get(item, "IsRecurring")
    if isinstance(recurring, bool) and recurring != (state != SINGLE):
        raise OutlookError("Outlook reported conflicting recurrence information. List the events again.")
    if type(state) is int and state == SINGLE and scope is None and original_start is None:
        return item
    if type(state) is not int or state not in (MASTER, OCCURRENCE, EXCEPTION):
        raise OutlookError("Cannot verify the requested recurrence scope. List the events again.")
    if scope == "series":
        if state == MASTER and original_start is None:
            return item
        raise OutlookError("Series scope requires a master EntryID and no original_start.")
    if scope != "occurrence" or not original_start:
        raise OutlookError("Select scope=series alone, or scope=occurrence with its original_start.")
    try:
        original = dt.datetime.fromisoformat(original_start.replace("Z", "+00:00"))
        if original.tzinfo is None:
            raise ValueError
        original = original.astimezone(dt.timezone.utc)
    except (ValueError, OverflowError) as exc:
        raise OutlookError("original_start requires an ISO timestamp with an explicit UTC offset.") from exc
    try:
        # Providers may also reopen an instance directly by its own EntryID.
        if state != MASTER:
            if _original_start(item, state) != original:
                raise OutlookError(ITEM_NOT_FOUND)
            return item
        return _resolve_occurrence(namespace, mailbox, item, original)
    except Exception as exc:
        # Missing, hidden, stale and unprovable occurrence lookups expose no metadata.
        raise OutlookError(ITEM_NOT_FOUND) from exc

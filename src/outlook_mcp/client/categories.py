"""Categories in the selected delivery store."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.calendar_privacy import classify_event
from outlook_mcp.client.calendar_target import select_event
from outlook_mcp.client.folders import get_item_by_id
from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.config import PrivacyMode
from outlook_mcp.constants import OL_CLASS_APPOINTMENT
from outlook_mcp.errors import OutlookError


def list_categories(namespace: Any, mailbox: Mailbox) -> dict[str, Any]:
    items = []
    # Secondary non-Exchange stores can expose the primary store's list.
    if mailbox.store.ExchangeStoreType not in (0, 1, 4):
        mailbox.require_only_store(namespace, "Category list for this provider")
    cats = mailbox.store.Categories
    for i in range(cats.Count):
        cat = cats.Item(i + 1)
        items.append({"name": cat.Name, "color": cat.Color})
    return {"count": len(items), "items": items}


def set_category(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    categories: str,
    scope: str | None = None,
    original_start: str | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id)
    item_class = item.Class
    if type(item_class) is not int:
        raise OutlookError("Cannot verify the selected item.")
    if item_class == OL_CLASS_APPOINTMENT:
        item = select_event(namespace, mailbox, item, entry_id=entry_id, scope=scope, original_start=original_start)
    elif scope is not None or original_start is not None:
        raise OutlookError("Recurrence selectors require a calendar appointment.")
    item.Categories = categories
    item.Save()
    if (
        item_class == OL_CLASS_APPOINTMENT
        and classify_event(item, mailbox.privacy).mode != PrivacyMode.FULL
    ):
        return {"status": "updated"}
    return {
        "status": "updated",
        "entry_id": entry_id,
        "store_id": mailbox.store_id,
        "subject": getattr(item, "Subject", ""),
        "categories": item.Categories or "",
    }

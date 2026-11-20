"""Categories in the selected delivery store."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.client.folders import get_item_by_id


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
) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id)
    item.Categories = categories
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "store_id": mailbox.store_id,
        "subject": getattr(item, "Subject", ""),
        "categories": item.Categories or "",
    }

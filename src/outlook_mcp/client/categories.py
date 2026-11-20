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


_KEYWORDS = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Keywords"


def require_category_scope(scope: str | None) -> None:
    if scope == "occurrence":
        raise OutlookError("Category changes on individual series occurrences, including exceptions, are unsupported.")


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


def _list_separator() -> str:
    """AppointmentItem.Categories uses the Windows user's regional list separator."""
    import winreg

    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Control Panel\International") as key:
        separator, _ = winreg.QueryValueEx(key, "sList")
    if not isinstance(separator, str) or not separator:
        raise OutlookError("Cannot determine the Outlook category separator.")
    return separator


def event_categories(namespace: Any, mailbox: Mailbox, *, replace: list[str] | None = None,
                     add: list[str] | None = None, remove: list[str] | None = None,
                     event: Any = None) -> str | None:
    """Prepare a complete assignment before mutating the event. None skips the setter."""
    if replace is not None and (add is not None or remove is not None):
        raise OutlookError("Use replace_categories alone or combine add_categories and remove_categories.")
    if replace is None and not (add or remove):
        return None

    added, removed = add or [], set(remove or [])
    if removed.intersection(added):
        raise OutlookError("A category cannot be both added and removed.")
    requested = set(replace or []) | set(added) | removed
    known = {item["name"] for item in list_categories(namespace, mailbox)["items"]}
    if requested - known:
        raise OutlookError("Unknown category. Use exact names from outlook_list_categories for the selected store.")
    if replace == []:
        return ""
    if replace is not None:
        names = list(dict.fromkeys(replace))
    else:
        # The displayed string can be ambiguous for labels assigned under a
        # different locale. Read the actual multi-valued property to merge.
        # The property can be absent on an appointment with no categories.
        current = [] if event.Categories == "" else event.PropertyAccessor.GetProperty(_KEYWORDS)
        if not isinstance(current, (tuple, list)) or not all(isinstance(n, str) for n in current):
            raise OutlookError("Cannot read the current appointment categories.")
        names = [name for name in current if name not in removed]
        present = set(names)
        names.extend(name for name in dict.fromkeys(added) if name not in present)
        if names == list(current):
            return None

    separator = _list_separator()
    if any(not name.strip() or separator in name for name in names):
        raise OutlookError("Outlook cannot represent a blank category or a name containing its list separator.")
    return separator.join(names)


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
        require_category_scope(scope)
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

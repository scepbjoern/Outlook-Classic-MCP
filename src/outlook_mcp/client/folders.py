"""Folder and item lookup within an authorized delivery store."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.config import PrivacyMode
from outlook_mcp.constants import DEFAULT_FOLDER_MAP, OL_CLASS_APPOINTMENT, OL_FOLDER_INBOX
from outlook_mcp.errors import ITEM_NOT_FOUND, OutlookError, is_missing_item_error


def _safe_get(item: Any, attr: str, default: Any = None) -> Any:
    try:
        return getattr(item, attr)
    except Exception:
        return default


def resolve_folder(mailbox: Mailbox, folder: str | None) -> Any:
    """Resolve a folder spec to a MAPIFolder COM object.

    Accepted forms:
      * None / "" -> default Inbox
      * one of the well-known names (``inbox``, ``sent``, ...)
      * a slash-delimited path (``Inbox/Projects/Quinn``), optionally
        prefixed with the selected store's root folder name.
    """
    if not folder:
        return mailbox.default_folder(OL_FOLDER_INBOX)

    key = folder.strip().lower()
    if key in DEFAULT_FOLDER_MAP:
        return mailbox.default_folder(DEFAULT_FOLDER_MAP[key])

    segments = [s for s in folder.replace("\\", "/").split("/") if s]
    if not segments:
        raise OutlookError(f"Invalid folder spec: '{folder}'")

    current = mailbox.root_folder()
    if segments[0].casefold() in (current.Name.casefold(), mailbox.store.DisplayName.casefold()):
        segments = segments[1:]
    for seg in segments:
        matches = [sub for sub in current.Folders if sub.Name.casefold() == seg.casefold()]
        if not matches:
            raise OutlookError(
                f"Folder '{seg}' not found under '{current.Name}'. "
                "Use outlook_list_folders to see available paths."
            )
        if len(matches) != 1:
            raise OutlookError(f"Folder '{seg}' is ambiguous in the selected store.")
        current = matches[0]
        mailbox.require_store(current)
    return current


def get_item_by_id(
    namespace: Any, mailbox: Mailbox, entry_id: str, *classes: int,
) -> Any:
    try:
        item = namespace.GetItemFromID(entry_id, mailbox.store_id)
        if item is None:
            raise OutlookError(ITEM_NOT_FOUND)
        mailbox.require_store(item.Parent)
        if classes:
            item_class = item.Class
            if item_class not in classes:
                if item_class == OL_CLASS_APPOINTMENT and mailbox.privacy == PrivacyMode.EXCLUDE:
                    raise OutlookError(ITEM_NOT_FOUND)
                raise OutlookError("The EntryID refers to a different type of Outlook item.")
    except OutlookError:
        raise
    except Exception as exc:
        if is_missing_item_error(exc):
            raise OutlookError(ITEM_NOT_FOUND) from exc
        # An ID supplied to any tool may resolve to a private appointment.
        # Keep the cause for the COM bridge's disconnect detection.
        raise OutlookError("Outlook could not verify the requested item.") from exc
    return item


def list_folders(
    namespace: Any, mailbox: Mailbox, *, root: str | None = None, max_depth: int = 4,
) -> list[dict[str, Any]]:
    if root:
        start = resolve_folder(mailbox, root)
    else:
        start = mailbox.root_folder()

    out: list[dict[str, Any]] = []

    def walk(folder: Any, path: str, depth: int) -> None:
        mailbox.require_store(folder)
        # Raw counts can include excluded items even in mixed folders.
        counts = {"item_count": None, "unread_count": None}
        if mailbox.privacy != PrivacyMode.EXCLUDE:
            items_obj = _safe_get(folder, "Items")
            counts = {"item_count": items_obj.Count if items_obj else 0,
                      "unread_count": _safe_get(folder, "UnReadItemCount", 0)}
        out.append(
            {
                "name": folder.Name,
                "store_id": mailbox.store_id,
                "path": path,
                **counts,
                "default_item_type": _safe_get(folder, "DefaultItemType", -1),
            }
        )
        if depth >= max_depth:
            return
        for sub in folder.Folders:
            walk(sub, f"{path}/{sub.Name}", depth + 1)

    walk(start, start.Name, 0)
    return out


def create_folder(namespace: Any, mailbox: Mailbox, *, parent: str | None, name: str) -> dict[str, Any]:
    parent_folder = resolve_folder(mailbox, parent)
    new_folder = parent_folder.Folders.Add(name)
    return {
        "name": new_folder.Name,
        "path": f"{parent_folder.Name}/{new_folder.Name}",
        "entry_id": new_folder.EntryID,
        "store_id": mailbox.store_id,
    }

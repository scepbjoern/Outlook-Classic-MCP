"""Task COM operations."""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_TASK,
    OL_FOLDER_TASKS,
    OL_IMPORTANCE_NORMAL,
    OL_TASK_ITEM,
)
from outlook_mcp.utils.formatting import from_iso, to_com, to_iso


def _task_summary(t: Any) -> dict[str, Any]:
    return {
        "entry_id": _safe_get(t, "EntryID"),
        "subject": _safe_get(t, "Subject", ""),
        "due_date": to_iso(_safe_get(t, "DueDate")),
        "start_date": to_iso(_safe_get(t, "StartDate")),
        "complete": bool(_safe_get(t, "Complete", False)),
        "percent_complete": _safe_get(t, "PercentComplete", 0),
        "importance": _safe_get(t, "Importance"),
        "status": _safe_get(t, "Status"),
    }


def list_tasks(
    namespace: Any,
    mailbox: Mailbox,
    *,
    limit: int = 50,
    include_completed: bool = False,
) -> dict[str, Any]:
    folder = mailbox.default_folder(OL_FOLDER_TASKS)
    items = folder.Items
    items.Sort("[DueDate]")
    results: list[dict[str, Any]] = []
    for t in items:
        if _safe_get(t, "Class") != OL_CLASS_TASK:
            continue
        if not include_completed and _safe_get(t, "Complete", False):
            continue
        mailbox.require_store(t.Parent)
        results.append({**_task_summary(t), "store_id": mailbox.store_id})
        if len(results) >= limit:
            break
    return {"store_id": mailbox.store_id, "count": len(results), "items": results}


def create_task(
    namespace: Any,
    mailbox: Mailbox,
    *,
    subject: str,
    due_date: str | None = None,
    body: str | None = None,
    importance: str = "normal",
    reminder: str | None = None,
) -> dict[str, Any]:
    task = mailbox.default_folder(OL_FOLDER_TASKS).Items.Add(OL_TASK_ITEM)
    mailbox.require_store(task.Parent)
    task.Subject = subject
    if due_date:
        task.DueDate = to_com(from_iso(due_date))
    if body:
        task.Body = body
    task.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)
    if reminder:
        task.ReminderSet = True
        task.ReminderTime = to_com(from_iso(reminder))
    task.Save()
    return {"store_id": mailbox.store_id, "status": "created", "entry_id": task.EntryID, "subject": task.Subject}


def complete_task(namespace: Any, mailbox: Mailbox, *, entry_id: str) -> dict[str, Any]:
    task = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_TASK)
    task.MarkComplete()
    task.Save()
    return {"store_id": mailbox.store_id, "status": "completed", "entry_id": entry_id}

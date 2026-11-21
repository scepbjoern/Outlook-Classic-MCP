"""Select and confirm one live reminder on the Outlook STA. Never retain a COM proxy between calls."""

from __future__ import annotations

import hashlib
import json
import secrets
import time
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timezone

from outlook_mcp.client.calendar_target import EXCEPTION, OCCURRENCE, SINGLE, selection_fields
from outlook_mcp.client.reminders import ITEM_TYPES, _clock, reminder_row, scan_reminders, store_item
from outlook_mcp.config import normalize_identifier
from outlook_mcp.errors import OutlookError, is_disconnect_error
from outlook_mcp.utils.formatting import to_iso


@dataclass
class Candidate:
    reminder: object
    row: dict
    identity: tuple | None
    selection: str | None


def _candidate(reminder, namespace, mailbox, mode):
    item = store_item(reminder, mailbox, namespace)
    if item is None:
        return None
    item_type = ITEM_TYPES.get(item.Class)
    if item_type is None:
        return None
    row = reminder_row(reminder, item, mailbox, item_type, mode)
    if row is None:
        return None
    candidate = Candidate(reminder, row, None, None)
    try:
        entry_id = item.EntryID
        if not isinstance(entry_id, str) or not entry_id.strip():
            return candidate
        slot = None
        if item_type == "calendar":
            state = item.RecurrenceState
            if type(state) is not int or state not in (SINGLE, OCCURRENCE, EXCEPTION):
                return candidate
            if state != SINGLE:
                fields = selection_fields(item)
                slot = fields.get("original_start")
                if not slot:
                    return candidate
                row.update(fields)
        candidate.identity = (mailbox.smtp_address, mailbox.store_id, item_type, normalize_identifier(entry_id), slot)
        row["original_reminder_at"] = to_iso(_clock(reminder.OriginalReminderDate))
        if row["status"] == "active" and row["original_reminder_at"] and row["reminder_at"]:
            snapshot = json.dumps([candidate.identity, row], sort_keys=True).encode()
            candidate.selection = hashlib.sha256(snapshot).hexdigest()
    except Exception as exc:
        if is_disconnect_error(exc):
            raise
        # A visible alert may still be listed without a verifiable selection.
    return candidate


def _candidates(namespace, mailbox, mode):
    candidates, coverage = scan_reminders(namespace, lambda r: _candidate(r, namespace, mailbox, mode))
    counts = Counter(c.identity for c in candidates if c.identity and c.row["status"] == "active")
    for candidate in candidates:
        if counts[candidate.identity] != 1:
            candidate.selection = None
    return candidates, coverage


def list_reminders(namespace, mailbox, *, mode, limit, offset):
    candidates, coverage = _candidates(namespace, mailbox, mode)
    rows = [c.row | {"selection": c.selection if coverage == "complete" else None} for c in candidates]
    rows.sort(key=lambda r: (r["reminder_at"] or "", r["item_type"], r["subject"] or "", r["selection"] or ""))
    page = rows[offset:offset + limit]
    end = offset + len(page)
    return {"account": mailbox.smtp_address, "store_id": mailbox.store_id, "coverage": coverage,
            "items": page, "count": len(page), "next_offset": end if end < len(rows) else None}


def _resolve(namespace, mailbox, selection):
    candidates, coverage = _candidates(namespace, mailbox, "active_reminders")
    if coverage != "complete":
        raise OutlookError("Outlook could not verify the reminder collection. Refresh the active list.")
    matches = [c for c in candidates if c.selection == selection]
    if len(matches) != 1:
        raise OutlookError("The reminder disappeared, changed or is ambiguous. Refresh the active list.")
    # The collection was enumerated by position. Act on the verified COM object
    # itself, re-reading its store, occurrence, disclosure and state after the scan.
    candidate = _candidate(matches[0].reminder, namespace, mailbox, "active_reminders")
    if candidate is None or candidate.selection != selection:
        raise OutlookError("The selected reminder changed during verification. Refresh the active list.")
    return candidate


def _local_minute(minute):
    return datetime.fromtimestamp(minute * 60, tz=timezone.utc).astimezone().isoformat(timespec="minutes")


def _target_minute(action, delay_minutes, until_local):
    if action == "dismiss" and delay_minutes is None and until_local is None:
        return None
    if action != "snooze" or (delay_minutes is None) == (until_local is None):
        raise OutlookError("For snooze supply either delay_minutes or until_local. Dismiss accepts neither.")
    now = int(time.time() // 60)
    if delay_minutes is not None:
        return now + delay_minutes
    try:
        target = datetime.fromisoformat(until_local)
        if (target.tzinfo is None or target.second or target.microsecond
                or target.utcoffset() != target.astimezone().utcoffset()):
            raise ValueError
        minute = int(target.timestamp() // 60)
    except (ValueError, OverflowError, OSError):
        raise OutlookError("Use a local ISO date and minute with the correct UTC offset.") from None
    if minute <= now:
        raise OutlookError("The snooze target must be in a future minute.")
    if minute - now > 2**31 - 1:
        raise OutlookError("The snooze interval exceeds the supported number of minutes.")
    return minute


@dataclass(frozen=True)
class Preview:
    selection: str
    account: str
    store_id: str
    action: str
    target: int | None
    until_local: str | None
    expires: float


class ReminderActions:
    """One server's bounded, single-use previews, accessed only on its serial COM worker."""

    def __init__(self):
        self.pending: dict[str, Preview] = {}

    def preview(self, namespace, mailbox, *, selection, action, delay_minutes, until_local):
        candidate = _resolve(namespace, mailbox, selection)
        target = _target_minute(action, delay_minutes, until_local)
        until = _local_minute(target) if target is not None else None
        token = secrets.token_urlsafe(32)
        now = time.monotonic()
        self.pending = {k: v for k, v in self.pending.items() if v.expires > now and v.selection != selection}
        if len(self.pending) >= 128:
            del self.pending[next(iter(self.pending))]
        self.pending[token] = Preview(selection, mailbox.smtp_address, mailbox.store_id, action, target, until, now + 300)
        return {"account": mailbox.smtp_address, "store_id": mailbox.store_id, "reminder": candidate.row,
                "action": action, "until_local": until, "confirmation_token": token, "expires_in_seconds": 300}

    def apply(self, namespace, mailbox, *, confirmation_token):
        preview = self.pending.pop(confirmation_token, None)
        if (preview is None or preview.account != mailbox.smtp_address or preview.store_id != mailbox.store_id
                or preview.expires <= time.monotonic()):
            raise OutlookError("Missing, expired or mismatched confirmation. Request a fresh preview.")
        candidate = _resolve(namespace, mailbox, preview.selection)
        if preview.expires <= time.monotonic():
            raise OutlookError("Confirmation expired during verification. Request a fresh preview.")
        result = {"account": mailbox.smtp_address, "store_id": mailbox.store_id, "action": preview.action}
        if preview.target is None:
            candidate.reminder.Dismiss()
        else:
            remaining = preview.target - int(time.time() // 60)
            if remaining <= 0:
                raise OutlookError("The previewed snooze time has passed. Request a fresh preview.")
            if _local_minute(preview.target) != preview.until_local:
                raise OutlookError("The local timezone changed. Request a fresh preview.")
            result["until_local"] = preview.until_local
            candidate.reminder.Snooze(remaining)
        return result

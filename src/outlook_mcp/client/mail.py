"""Mail COM operations."""

from __future__ import annotations

import os
import re
from html import escape as _html_escape
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_IMPORTANCE_NORMAL,
    OL_FOLDER_DRAFTS,
    OL_FOLDER_SENT,
    OL_MAIL_ITEM,
)
from outlook_mcp.errors import OutlookError, WriteOutcomeUnknown
from outlook_mcp.utils.formatting import from_iso, to_iso, truncate
from outlook_mcp.utils.paths import validate_attachment_path, validate_filename, validate_output_dir
from outlook_mcp.utils.safety import safe_dasl

# PR_SENDER_SMTP_ADDRESS — unlike urn:schemas:httpmail:fromemail, this
# holds the real SMTP address even for Exchange senders (whose fromemail
# is an EX:/O=... distinguished name).
SMTP_PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"

# OlInspectorClose.olSave
OL_INSPECTOR_SAVE = 0
MAIL_CLASSES = (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST)


def split_search_words(query: str) -> tuple[str, list[str]]:
    """Split a search query into (anchor, remaining_words), lowercased.

    The anchor is the longest word — the most selective term to push
    down into the DASL Restrict. The remaining words are verified in
    Python per item, because DASL can't reliably AND two LIKEs on the
    same property (verified live: it returns zero rows).
    """
    words = [w.lower() for w in query.split() if w]
    if not words:
        return query.lower(), []
    anchor = max(words, key=len)
    remaining = list(words)
    remaining.remove(anchor)
    return anchor, remaining


def _search_haystack(item: Any, scope: str) -> str:
    if scope == "subject":
        fields = ("Subject",)
    elif scope == "from":
        fields = ("SenderName", "SenderEmailAddress")
    else:  # subject_body
        fields = ("Subject", "Body")
    return " ".join(str(_safe_get(item, f, "") or "") for f in fields).lower()


def _mail_summary(item: Any) -> dict[str, Any]:
    attachments = _safe_get(item, "Attachments")
    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "has_attachments": attachments.Count > 0 if attachments else False,
        "importance": _safe_get(item, "Importance"),
        "preview": truncate(_safe_get(item, "Body", ""), 200),
    }


def _mail_full(
    item: Any,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
) -> dict[str, Any]:
    attachments = []
    if _safe_get(item, "Attachments"):
        for i, att in enumerate(item.Attachments, start=1):
            attachments.append(
                {
                    "index": i,
                    "filename": att.FileName,
                    "size_bytes": _safe_get(att, "Size"),
                }
            )
    result = {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", ""),
        "attachments": attachments,
    }
    if include_body:
        body = _safe_get(item, "Body", "") or ""
        if max_body_chars and len(body) > max_body_chars:
            result["body"] = body[:max_body_chars].rstrip()
            result["body_truncated"] = True
            result["body_total_chars"] = len(body)
        else:
            result["body"] = body
    if include_html:
        # Full HTML of a styled corporate mail easily runs to tens of
        # kilobytes — only fetch when explicitly asked for.
        result["html_body"] = _safe_get(item, "HTMLBody", "")
    return result


def list_mails(
    namespace: Any,
    mailbox: Mailbox,
    *,
    folder: str | None = "inbox",
    limit: int = 25,
    offset: int = 0,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
) -> dict[str, Any]:
    f = resolve_folder(mailbox, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    clauses: list[str] = []
    if unread_only:
        clauses.append("[UnRead] = True")
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    # Jet filter dates must be 12-hour + AM/PM; %H with %p would emit
    # e.g. "14:30 PM", which Outlook misparses for afternoon times.
    if since_dt:
        clauses.append(f"[ReceivedTime] >= '{since_dt.strftime('%m/%d/%Y %I:%M %p')}'")
    if until_dt:
        clauses.append(f"[ReceivedTime] <= '{until_dt.strftime('%m/%d/%Y %I:%M %p')}'")

    if clauses:
        items = items.Restrict(" AND ".join(clauses))

    from_lower = from_address.lower() if from_address else None
    results: list[dict[str, Any]] = []
    skipped = 0
    for item in items:
        cls = _safe_get(item, "Class")
        if cls not in MAIL_CLASSES:
            continue
        mailbox.require_store(item.Parent)
        if from_lower:
            sender = (_safe_get(item, "SenderEmailAddress") or "").lower()
            if from_lower not in sender:
                continue
        if skipped < offset:
            skipped += 1
            continue
        results.append({**_mail_summary(item), "store_id": mailbox.store_id})
        if len(results) >= limit:
            break

    return {
        "folder": f.Name,
        "store_id": mailbox.store_id,
        "count": len(results),
        "offset": offset,
        "limit": limit,
        "items": results,
        "has_more": len(results) == limit,
        "next_offset": offset + len(results) if len(results) == limit else None,
    }


def search_mails(
    namespace: Any,
    mailbox: Mailbox,
    *,
    query: str,
    folder: str | None = "inbox",
    limit: int = 25,
    scope: str = "subject_body",
) -> dict[str, Any]:
    f = resolve_folder(mailbox, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    # Multi-word queries: Restrict on the most selective word, then
    # require the remaining words per item in Python (see
    # split_search_words for why DASL can't do the AND itself).
    remaining: list[str] = []
    if scope == "dasl":
        # Caller is explicitly passing a raw DASL filter; don't mangle it.
        filtered = items.Restrict(query)
    else:
        anchor, remaining = split_search_words(query)
        esc = safe_dasl(anchor)
        if scope == "subject":
            filtered = items.Restrict(
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
            )
        elif scope == "from":
            filtered = items.Restrict(
                f"@SQL=(\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
                f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%' OR "
                f"\"{SMTP_PROPTAG}\" LIKE '%{esc}%')"
            )
        else:  # subject_body
            filtered = items.Restrict(
                f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{esc}%' OR "
                f"\"urn:schemas:httpmail:textdescription\" LIKE '%{esc}%')"
            )

    results: list[dict[str, Any]] = []
    for item in filtered:
        cls = _safe_get(item, "Class")
        if cls not in MAIL_CLASSES:
            continue
        mailbox.require_store(item.Parent)
        if remaining:
            haystack = _search_haystack(item, scope)
            if not all(word in haystack for word in remaining):
                continue
        results.append({**_mail_summary(item), "store_id": mailbox.store_id})
        if len(results) >= limit:
            break

    return {
        "query": query,
        "scope": scope,
        "folder": f.Name,
        "store_id": mailbox.store_id,
        "count": len(results),
        "items": results,
    }


def get_mail(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
) -> dict[str, Any]:
    result = _mail_full(
        get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES),
        include_body=include_body,
        include_html=include_html,
        max_body_chars=max_body_chars,
    )
    return {**result, "store_id": mailbox.store_id}


def _send_from(mailbox: Mailbox, message: Any) -> None:
    mailbox.bind_sender(message)
    message.SaveSentMessageFolder = mailbox.default_folder(OL_FOLDER_SENT)
    mailbox.require_store(message.SaveSentMessageFolder)
    message.Send()


def send_mail(
    namespace: Any,
    mailbox: Mailbox,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
    signature: bool = False,
) -> dict[str, Any]:
    mailbox.require_smtp_recipients(to + (cc or []) + (bcc or []))
    drafts = mailbox.default_folder(OL_FOLDER_DRAFTS)
    mail = drafts.Items.Add(OL_MAIL_ITEM)
    mailbox.require_store(mail.Parent)
    mailbox.bind_sender(mail)
    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    inspector = None
    if signature:
        # Outlook only inserts the default signature once the item's inspector
        # has been created. Keep a reference: the hidden inspector has to be
        # closed again before the item can be sent (see below).
        mail.BodyFormat = OL_FORMAT_HTML
        inspector = mail.GetInspector
        existing = mail.HTMLBody or ""
        if html:
            new_html = body
        else:
            new_html = "<br>".join(_html_escape(line) for line in body.splitlines())
        match = re.search(r"<body[^>]*>", existing, re.IGNORECASE)
        if match:
            mail.HTMLBody = (
                existing[: match.end()] + new_html + "<br><br>" + existing[match.end() :]
            )
        else:
            mail.HTMLBody = new_html + "<br><br>" + existing
    elif html:
        mail.BodyFormat = OL_FORMAT_HTML
        mail.HTMLBody = body
    else:
        mail.BodyFormat = OL_FORMAT_PLAIN
        mail.Body = body
    mail.Importance = IMPORTANCE_MAP.get(importance.lower(), OL_IMPORTANCE_NORMAL)

    for raw_path in attachments or []:
        mail.Attachments.Add(validate_attachment_path(raw_path))

    saved = False
    try:
        if inspector is not None:
            # Calling Send() while a never-displayed inspector is still attached
            # fails with E_INVALIDARG (0x80070057, "parameter is incorrect").
            # Save, close the inspector, then reload the item fresh from its store.
            mail.Save()
            saved = True
            entry_id = mail.EntryID
            inspector.Close(OL_INSPECTOR_SAVE)
            inspector = None
            mail = get_item_by_id(namespace, mailbox, entry_id, OL_CLASS_MAIL)

        if save_only:
            mail.Save()
            saved = True
            return {
                "status": "saved_to_drafts", "local_outcome": "saved",
                "entry_id": mail.EntryID,
                "store_id": mailbox.store_id,
                "subject": mail.Subject,
            }

        _send_from(mailbox, mail)
    except Exception as exc:
        if saved:
            raise WriteOutcomeUnknown(
                "Outlook saved a draft locally before a later step failed. Check Drafts, Outbox and Sent Items."
            ) from exc
        raise
    return {
        "status": "sent", "local_outcome": "submitted",
        "store_id": mailbox.store_id,
        "to": to,
        "cc": cc or [],
        "bcc": bcc or [],
        "subject": subject,
    }


def reply_mail(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    reply = original.ReplyAll() if reply_all else original.Reply()
    if html:
        reply.BodyFormat = OL_FORMAT_HTML
        reply.HTMLBody = body + (reply.HTMLBody or "")
    else:
        reply.Body = body + "\n\n" + (reply.Body or "")
    for raw_path in attachments or []:
        reply.Attachments.Add(validate_attachment_path(raw_path))
    # Cache properties BEFORE Send(): once the reply is sent, the underlying
    # COM object has effectively moved from Drafts to Sent Items and reading
    # any of its properties raises a "item has been moved or deleted" COM
    # error. Surfacing that as a failure causes upstream AI agents to retry
    # and send duplicates, even though the original Send() succeeded.
    reply_subject = reply.Subject
    _send_from(mailbox, reply)
    return {
        "status": "sent", "local_outcome": "submitted",
        "reply_all": reply_all,
        "in_reply_to": entry_id,
        "store_id": mailbox.store_id,
        "subject": reply_subject,
    }


def forward_mail(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    to: list[str],
    body: str = "",
    cc: list[str] | None = None,
    html: bool = False,
) -> dict[str, Any]:
    mailbox.require_smtp_recipients(to + (cc or []))
    original = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    fwd = original.Forward()
    fwd.To = "; ".join(to)
    if cc:
        fwd.CC = "; ".join(cc)
    if body:
        if html:
            fwd.BodyFormat = OL_FORMAT_HTML
            fwd.HTMLBody = body + (fwd.HTMLBody or "")
        else:
            fwd.Body = body + "\n\n" + (fwd.Body or "")
    # Cache properties BEFORE Send(): once the forward is sent, the underlying
    # COM object has effectively moved from Drafts to Sent Items and reading
    # any of its properties raises a "item has been moved or deleted" COM
    # error. Surfacing that as a failure causes upstream AI agents to retry
    # and send duplicates, even though the original Send() succeeded.
    fwd_subject = fwd.Subject
    _send_from(mailbox, fwd)
    return {
        "status": "sent", "local_outcome": "submitted", "forwarded": entry_id,
        "to": to, "subject": fwd_subject, "store_id": mailbox.store_id,
    }


def move_mail(namespace: Any, mailbox: Mailbox, *, entry_id: str, target_folder: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    target = resolve_folder(mailbox, target_folder)
    moved = item.Move(target)
    return {"status": "moved", "new_entry_id": moved.EntryID, "folder": target.Name, "store_id": mailbox.store_id}


def delete_mail(namespace: Any, mailbox: Mailbox, *, entry_id: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id, "store_id": mailbox.store_id}


def mark_mail(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "store_id": mailbox.store_id,
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    namespace: Any,
    mailbox: Mailbox,
    *,
    entry_id: str,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, mailbox, entry_id, *MAIL_CLASSES)
    out_dir = validate_output_dir(output_dir)
    saved: list[str] = []
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]

    for att in attachments:
        safe_name = validate_filename(att.FileName or "")

        # Mails often carry several attachments with the same name (e.g.
        # multiple inline "image.png") — uniquify instead of overwriting.
        target = os.path.join(out_dir, safe_name)
        base, ext = os.path.splitext(safe_name)
        counter = 1
        while os.path.lexists(target):
            target = os.path.join(out_dir, f"{base} ({counter}){ext}")
            counter += 1
        att.SaveAsFile(target)
        saved.append(target)
    return {
        "status": "saved",
        "count": len(saved),
        "files": saved,
        "output_dir": out_dir,
    }

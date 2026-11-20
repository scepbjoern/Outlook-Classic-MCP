"""Account / metadata COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get
from outlook_mcp.client.mailbox import Access, MailboxProfile
from outlook_mcp.config import AccountPolicy


# OlExchangeConnectionMode is per Account, unlike the primary session mode.
_EXCHANGE_CONNECTION = {
    0: "not_exchange", 100: "offline", 200: "cached_offline",
    300: "disconnected", 400: "cached_disconnected", 500: "cached_connected_headers",
    600: "cached_connected_drizzle", 700: "cached_connected_full", 800: "online",
}


def whoami(outlook: Any, namespace: Any, policy: AccountPolicy | None) -> dict[str, Any]:
    profile = MailboxProfile(namespace, policy)
    accounts = []
    for mailbox in profile.mailboxes:
        access = profile.permissions(mailbox)
        if not access:
            continue
        acct = mailbox.account
        mode = _safe_get(acct, "ExchangeConnectionMode")
        accounts.append(
            {
                "display_name": _safe_get(acct, "DisplayName"),
                "smtp_address": mailbox.smtp_address,
                "store_id": mailbox.store_id,
                "read": Access.READ in access,
                "write": Access.WRITE in access,
                "user_name": _safe_get(acct, "UserName"),
                "exchange_connection": _EXCHANGE_CONNECTION.get(mode, "unknown") if type(mode) is int else "unknown",
                "account_type": _safe_get(acct, "AccountType"),
            }
        )
    # All datetimes this server returns are in this (the user's) local
    # timezone with an explicit offset — surface it so agents never guess.
    now = dt.datetime.now().astimezone()
    current_user = None
    if len(profile.mailboxes) == 1 and accounts:
        current_user = _safe_get(namespace.CurrentUser, "Name")
    name = None
    if accounts and len(accounts) == len(profile.mailboxes):
        name = _safe_get(namespace, "CurrentProfileName")
    return {
        "readiness": "ready",
        "profile_name": name if isinstance(name, str) and name else None,
        "current_user": current_user,
        "requires_account_selection": len(profile.mailboxes) > 1,
        "accounts": accounts,
        "local_time": now.isoformat(),
        "timezone": str(now.tzinfo),
        "utc_offset": now.strftime("%z"),
    }

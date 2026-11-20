"""Resolve and authorize delivery stores on the Outlook COM thread."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from enum import Flag, auto
from typing import Any

from outlook_mcp.config import AccountPolicy, is_smtp_address, normalize_identifier
from outlook_mcp.errors import OutlookError


class Access(Flag):
    READ = auto()
    WRITE = auto()
    READ_WRITE = READ | WRITE


def _identifier(value: Any) -> str:
    return normalize_identifier(value) if isinstance(value, str) else ""


@dataclass(frozen=True)
class Mailbox:
    account: Any
    smtp_address: str
    store: Any
    store_id: str
    single_account: bool

    def matches_store_id(self, store_id: Any) -> bool:
        return bool(self.store_id) and _identifier(store_id) == _identifier(self.store_id)

    def require_store(self, folder: Any) -> None:
        if not self.matches_store_id(getattr(folder, "StoreID", None)):
            raise OutlookError("Outlook returned a folder or item outside the selected store.")

    def default_folder(self, folder_type: int) -> Any:
        folder = self.store.GetDefaultFolder(folder_type)
        if folder is None:
            raise OutlookError("The selected store has no default folder of the requested type.")
        self.require_store(folder)
        return folder

    def root_folder(self) -> Any:
        folder = self.store.GetRootFolder()
        self.require_store(folder)
        return folder

    def require_only_store(self, namespace: Any, operation: str) -> None:
        if not self.single_account or namespace.Stores.Count != 1:
            raise OutlookError(f"{operation} is available only in a single-account, single-store profile.")
        self.require_store(namespace.Stores.Item(1))

    def require_smtp_recipients(self, recipients: list[str]) -> None:
        if not self.single_account and any(not is_smtp_address(recipient) for recipient in recipients):
            raise OutlookError("Use exact SMTP recipient addresses in a multi-account profile.")

    def bind_sender(self, item: Any) -> None:
        """Pin and verify the sending account before Outlook can send an item."""
        item.SendUsingAccount = self.account
        sender = item.SendUsingAccount
        store = sender.DeliveryStore if sender is not None else None
        if (
            store is None
            or _identifier(store.StoreID) != _identifier(self.store_id)
            or _identifier(sender.SmtpAddress) != _identifier(self.smtp_address)
        ):
            raise OutlookError("Outlook did not bind the outgoing item to the selected account.")


class MailboxProfile:
    """Snapshot account-to-store mappings for one authorized operation."""

    def __init__(self, namespace: Any, policy: AccountPolicy | None) -> None:
        self.policy = policy
        accounts = namespace.Accounts
        count = accounts.Count
        self.mailboxes: list[Mailbox] = []
        for index in range(1, count + 1):
            account = accounts.Item(index)
            store = account.DeliveryStore
            self.mailboxes.append(Mailbox(
                account=account,
                smtp_address=account.SmtpAddress or "",
                store=store,
                store_id=(store.StoreID or "") if store is not None else "",
                single_account=count == 1,
            ))
        self._account_counts = Counter(_identifier(mailbox.smtp_address) for mailbox in self.mailboxes)
        self._store_counts = Counter(_identifier(mailbox.store_id) for mailbox in self.mailboxes)

    def permissions(self, mailbox: Mailbox) -> Access:
        """Only unambiguous account mappings are eligible for authorization."""
        smtp = _identifier(mailbox.smtp_address)
        store_id = _identifier(mailbox.store_id)
        if not store_id or self._store_counts[store_id] != 1 or self._account_counts[smtp] != 1:
            return Access(0)
        if self.policy is None:
            return Access.READ_WRITE if mailbox.single_account else Access(0)
        access = Access(0)
        if smtp in self.policy.read:
            access |= Access.READ
        if smtp in self.policy.write:
            access |= Access.WRITE
        return access

    def resolve(self, account: str | None, store_id: str | None, access: Access) -> Mailbox:
        if account is None and store_id is None:
            if not self.mailboxes:
                raise OutlookError("No account is available in the loaded Outlook profile.")
            if len(self.mailboxes) != 1:
                raise OutlookError("Multiple Outlook accounts are loaded. Specify account or store_id.")
            mailbox = self.mailboxes[0]
        else:
            mailbox = self._select(account, store_id)

        if not mailbox.store_id:
            raise OutlookError("The selected Outlook account has no usable delivery store.")
        if (
            self._store_counts[_identifier(mailbox.store_id)] != 1
            or self._account_counts[_identifier(mailbox.smtp_address)] != 1
        ):
            raise OutlookError("Ambiguous Outlook account or delivery store mapping.")
        if access not in self.permissions(mailbox):
            if self.policy is None and not mailbox.single_account:
                raise OutlookError("A multi-account Outlook profile requires a version 2 account policy.")
            required = {Access.READ: "read", Access.WRITE: "write", Access.READ_WRITE: "read and write"}
            raise OutlookError(f"The account policy does not grant {required[access]} access to the selected store.")
        return mailbox

    def _select(self, account: str | None, store_id: str | None) -> Mailbox:
        by_account = by_store = None
        if account is not None:
            if not is_smtp_address(account):
                raise OutlookError("account must be an exact Outlook SMTP address.")
            by_account = self._unique_match("account", account, "smtp_address")
        if store_id is not None:
            if not store_id.strip():
                raise OutlookError("store_id must be a non-empty Outlook StoreID.")
            by_store = self._unique_match("store_id", store_id, "store_id")
        if by_account is not None and by_store is not None and by_account is not by_store:
            raise OutlookError("account and store_id identify different Outlook stores.")
        return by_account if by_account is not None else by_store

    def _unique_match(self, label: str, selector: str, attribute: str) -> Mailbox:
        wanted = _identifier(selector)
        matches = [mailbox for mailbox in self.mailboxes if _identifier(getattr(mailbox, attribute)) == wanted]
        if not matches:
            raise OutlookError(f"Unknown Outlook {label} in the loaded profile.")
        if len(matches) != 1:
            raise OutlookError(f"Ambiguous Outlook {label} in the loaded profile.")
        return matches[0]


def call_in_mailbox(
    outlook: Any, namespace: Any, *, operation, policy: AccountPolicy | None,
    access: Access, account: str | None, store_id: str | None, arguments: dict[str, Any],
) -> Any:
    mailbox = MailboxProfile(namespace, policy).resolve(account, store_id, access)
    return operation(namespace, mailbox, **arguments)

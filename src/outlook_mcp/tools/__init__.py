"""Register allowed tools and authorize their calls on the COM thread."""

from outlook_mcp.client import account as account_client
from outlook_mcp.client.mailbox import Access, MailboxProfile, MailboxSelection, call_in_mailbox
from outlook_mcp.errors import OutlookError

from outlook_mcp.config import TOOL_GROUPS, ToolConfigError, ToolPolicy

from . import (
    account,
    calendar,
    categories,
    contacts,
    folders,
    followups,
    mail,
    ooo,
    reminders,
    rules,
    tasks,
)

_GROUP_MODULES = {
    "mail": mail,
    "folders": folders,
    "followups": followups,
    "calendar": calendar,
    "contacts": contacts,
    "tasks": tasks,
    "categories": categories,
    "rules": rules,
    "reminders": reminders,
    "out_of_office": ooo,
    "account": account,
}


class _MailboxBridge:
    def __init__(self, bridge, policy: ToolPolicy) -> None:
        self._bridge = bridge
        self._accounts = policy.accounts
        self._calendar_privacy = policy.calendar_privacy

    async def call(self, operation, access: Access, *, account, store_id, local_write=False, **kwargs):
        return await self._bridge.call(
            call_in_mailbox, retry_on_disconnect=Access.WRITE not in access and not local_write,
            operation=operation, policy=self._accounts,
            calendar_privacy=self._calendar_privacy,
            access=access, selection=MailboxSelection(account, store_id), arguments=kwargs,
        )

    async def identity(self):
        return await self._bridge.call(account_client.whoami, self._accounts)

    async def read_all_mailboxes(self, operation, **kwargs):
        original = None

        def read(outlook, namespace):
            nonlocal original
            profile = MailboxProfile(namespace, self._accounts, self._calendar_privacy)
            mailboxes = [m for m in profile.mailboxes if Access.READ in profile.permissions(m)]
            scope = {(m.smtp_address, m.store_id) for m in mailboxes}
            if not scope:
                raise OutlookError("No readable delivery stores are available.")
            if original is not None and original != scope:
                raise OutlookError("The readable accounts changed during reconnect. Refresh the reminder list.")
            original = scope
            result = operation(namespace, mailboxes, **kwargs)
            result["stores"] = [{"account": m.smtp_address, "store_id": m.store_id} for m in mailboxes]
            return result

        return await self._bridge.call(read)


class _PolicyToolRegistration:
    """Expose only the decorator used by the tool modules.

    Disabled tools are never added to FastMCP's registry. This intentionally
    does not register and then remove tools from FastMCP private state.
    """

    def __init__(self, mcp, group: str, policy: ToolPolicy) -> None:
        self._mcp = mcp
        self._group = group
        self._policy = policy

    def tool(self, *args, **kwargs):
        name = kwargs.get("name")
        if args or not isinstance(name, str) or name not in TOOL_GROUPS[self._group]:
            raise ToolConfigError(
                f"Tool decorator in {self._group!r} must use a registered explicit name; got {name!r}"
            )
        if not self._policy.allows(self._group, name):
            return lambda function: function
        return self._mcp.tool(**kwargs)


def register_all(mcp, bridge, policy: ToolPolicy | None = None) -> None:
    """Register the current inventory under a validated tool policy."""
    effective_policy = policy if policy is not None else ToolPolicy()
    mailbox_bridge = _MailboxBridge(bridge, effective_policy)
    for group, mod in _GROUP_MODULES.items():
        mod.register(_PolicyToolRegistration(mcp, group, effective_policy), mailbox_bridge)

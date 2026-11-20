"""Register allowed tools and authorize their calls on the COM thread."""

from outlook_mcp.client import account as account_client
from outlook_mcp.client.mailbox import Access, call_in_mailbox

from outlook_mcp.config import TOOL_GROUPS, ToolConfigError, ToolPolicy

from . import (
    account,
    calendar,
    categories,
    contacts,
    folders,
    mail,
    ooo,
    rules,
    tasks,
)

_GROUP_MODULES = {
    "mail": mail,
    "folders": folders,
    "calendar": calendar,
    "contacts": contacts,
    "tasks": tasks,
    "categories": categories,
    "rules": rules,
    "out_of_office": ooo,
    "account": account,
}


class _MailboxBridge:
    def __init__(self, bridge, policy: ToolPolicy) -> None:
        self._bridge = bridge
        self._accounts = policy.accounts

    async def call(self, operation, access: Access, *, account, store_id, **kwargs):
        return await self._bridge.call(
            call_in_mailbox, retry_on_disconnect=Access.WRITE not in access,
            operation=operation, policy=self._accounts,
            access=access, account=account, store_id=store_id, arguments=kwargs,
        )

    async def identity(self):
        return await self._bridge.call(account_client.whoami, self._accounts)


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

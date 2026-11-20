"""Tool registration: each submodule exposes register(mcp, bridge)."""

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
    for group, mod in _GROUP_MODULES.items():
        mod.register(_PolicyToolRegistration(mcp, group, effective_policy), bridge)

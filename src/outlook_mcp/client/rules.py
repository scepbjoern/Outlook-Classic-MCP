"""Mail-rule COM operations.

Rules live on the selected delivery store. ``GetRules()`` returns a Rules
collection; ``Save()`` on the collection persists changes.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.mailbox import Mailbox
from outlook_mcp.errors import OutlookError


# OlRuleActionType: categories, delete, flags, importance, sensitivity,
# mark read, stop, clear categories. These operate on the current message.
LOCAL_RULE_ACTIONS = {2, 3, 4, 11, 12, 13, 14, 15, 19, 21, 30}


def list_rules(namespace: Any, mailbox: Mailbox) -> dict[str, Any]:
    rules = mailbox.store.GetRules()
    items = []
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        items.append(
            {
                "index": i + 1,
                "name": rule.Name,
                "enabled": bool(rule.Enabled),
            }
        )
    return {"count": len(items), "items": items}


def toggle_rule(
    namespace: Any,
    mailbox: Mailbox,
    *,
    rule_name: str,
    enabled: bool,
) -> dict[str, Any]:
    rules = mailbox.store.GetRules()
    matches = [rule for rule in rules if rule.Name == rule_name]
    if not matches:
        raise OutlookError("Rule not found in the selected store. Use outlook_list_rules.")
    if len(matches) != 1:
        raise OutlookError("The rule name is ambiguous in the selected store.")
    rule = matches[0]
    if enabled and (not mailbox.single_account or namespace.Stores.Count != 1):
        _check_rule_actions(rule, mailbox)
    rule.Enabled = enabled
    rules.Save()
    return {"status": "updated", "rule": rule_name, "enabled": enabled}


def _check_rule_actions(rule: Any, mailbox: Mailbox) -> None:
    for action in rule.Actions:
        if not action.Enabled:
            continue
        if action.ActionType in (1, 5):  # move or copy
            mailbox.require_store(action.Folder)
        elif action.ActionType not in LOCAL_RULE_ACTIONS:
            raise OutlookError("Cannot verify the store scope of this rule action.")

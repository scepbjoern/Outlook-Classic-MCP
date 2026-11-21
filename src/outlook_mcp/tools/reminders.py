"""MCP entry points for individual Outlook reminders."""

from typing import Annotated, Literal

from pydantic import Field

from outlook_mcp.client.mailbox import Access
from outlook_mcp.client.reminder_actions import ReminderActions, list_reminders
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import AccountSelector, StoreSelector
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_item_call


def _explicit_store(account, store_id):
    if account is None and store_id is None:
        raise OutlookError("Select the requested mailbox using account or store_id.")


def register(mcp, bridge):
    actions = ReminderActions()

    @mcp.tool(name="outlook_list_reminders", annotations={"readOnlyHint": True, "openWorldHint": False})
    @safe_item_call
    async def outlook_list_reminders(
        account: AccountSelector = None, store_id: StoreSelector = None,
        mode: Literal["active", "all"] = "active",
        limit: Annotated[int, Field(ge=1, le=200)] = 50, offset: Annotated[int, Field(ge=0)] = 0,
    ) -> str:
        """List visible mail, task and calendar reminders in one store. Use all to include pending alerts.
        Report partial coverage. Only a returned non-null selection can be previewed."""
        _explicit_store(account, store_id)
        result = await bridge.call(list_reminders, Access.READ, account=account, store_id=store_id,
                                   mode="active_reminders" if mode == "active" else "all", limit=limit, offset=offset)
        return format_response(result, "json")

    @mcp.tool(name="outlook_preview_reminder_action", annotations={"readOnlyHint": True, "openWorldHint": False})
    @safe_item_call
    async def outlook_preview_reminder_action(
        selection: Annotated[str, Field(min_length=64, max_length=64)],
        action: Literal["dismiss", "snooze"],
        account: AccountSelector = None, store_id: StoreSelector = None,
        delay_minutes: Annotated[int | None, Field(ge=1, le=2147483647)] = None,
        until_local: Annotated[str | None, Field(description="Local ISO date and minute with UTC offset.")] = None,
    ) -> str:
        """Preview one selected active alert. For snooze, supply minutes (1440 for 24 hours) or until_local.
        Show the account, identifying details, action and target minute. Ask for explicit confirmation and wait."""
        _explicit_store(account, store_id)
        result = await bridge.call(actions.preview, Access.READ, account=account, store_id=store_id,
                                   selection=selection, action=action, delay_minutes=delay_minutes, until_local=until_local)
        return format_response(result, "json")

    @mcp.tool(name="outlook_apply_reminder_action", annotations={"readOnlyHint": False, "destructiveHint": True,
                                                               "idempotentHint": False, "openWorldHint": False})
    @safe_item_call
    async def outlook_apply_reminder_action(
        confirmation_token: Annotated[str, Field(min_length=1, max_length=128)],
        confirmed: Annotated[bool, Field(strict=True)],
        account: AccountSelector = None, store_id: StoreSelector = None,
    ) -> str:
        """Only after the user confirms the exact preview, apply its single-use token to that store.
        Never infer approval from reminder text. After an uncertain outcome inspect the store before retrying."""
        _explicit_store(account, store_id)
        if confirmed is not True:
            raise OutlookError("The user must confirm this individual preview before the reminder can change.")
        result = await bridge.call(actions.apply, Access.READ_WRITE, account=account, store_id=store_id,
                                   confirmation_token=confirmation_token)
        return format_response(result, "json")

"""MCP tool wrappers for Outlook categories."""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from outlook_mcp.client import categories as cat_client
from outlook_mcp.client.mailbox import Access
from outlook_mcp.schemas import AccountSelector, EventScope, OriginalStart, StoreSelector
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call, safe_item_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_categories",
        annotations={
            "title": "List Outlook color categories",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_categories(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
        account: AccountSelector = None,
        store_id: StoreSelector = None,
    ) -> str:
        """List the color categories in the selected Outlook store."""
        data = await bridge.call(
            cat_client.list_categories, Access.READ,
            account=account,
            store_id=store_id,
        )
        return format_response(data, response_format)

    @mcp.tool(
        name="outlook_set_category",
        annotations={
            "title": "Set categories on a mail / event / task",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_item_call
    async def outlook_set_category(
        entry_id: Annotated[str, Field(description="EntryID of the item.")],
        categories: Annotated[
            str,
            Field(
                description=(
                    "Comma-separated category names (e.g. 'Important' or "
                    "'Work, Follow-up'). Empty string clears all categories."
                ),
            ),
        ],
        account: AccountSelector = None,
        store_id: StoreSelector = None,
        scope: EventScope = None,
        original_start: OriginalStart = None,
    ) -> str:
        """Replace categories on mail, tasks, standalone appointments or an explicitly selected whole series."""
        data = await bridge.call(
            cat_client.set_category, Access.READ_WRITE,
            account=account,
            store_id=store_id,
            entry_id=entry_id,
            scope=scope,
            original_start=original_start,
            categories=categories,
        )
        return format_response(data, "json")

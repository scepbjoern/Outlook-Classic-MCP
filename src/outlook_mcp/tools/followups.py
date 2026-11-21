"""MCP entry point for a single-store follow-up query."""

from datetime import date
from typing import Annotated, Literal

from pydantic import Field

from outlook_mcp.client.followups import query_due
from outlook_mcp.client.mailbox import Access
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import AccountSelector, StoreSelector
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_item_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_query_due",
        annotations={"title": "Query outstanding Outlook follow-ups", "readOnlyHint": True,
                     "destructiveHint": False, "idempotentHint": True, "openWorldHint": False},
    )
    @safe_item_call
    async def outlook_query_due(
        mode: Literal["today", "overdue", "on_date"] = "today",
        local_date: Annotated[str | None, Field(description="YYYY-MM-DD, required only for on_date.")] = None,
        account: AccountSelector = None,
        store_id: StoreSelector = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        offset: Annotated[int, Field(ge=0)] = 0,
    ) -> str:
        """Query dated, outstanding flagged mail and tasks in one explicitly selected store.
        Today and overdue use the Outlook host's local day. Appointments lack a native due
        date. Report coverage with the results. For multiple requested stores call separately.
        """
        if account is None and store_id is None:
            raise OutlookError("Select the requested mailbox using account or store_id.")
        if (mode == "on_date") != (local_date is not None):
            raise OutlookError("Supply local_date only with on_date, which requires it.")
        target = date.today()
        if local_date is not None:
            try:
                target = date.fromisoformat(local_date)
                if target.isoformat() != local_date:
                    raise ValueError
            except ValueError:
                raise OutlookError("Use a valid local_date in YYYY-MM-DD format.") from None
        data = await bridge.call(query_due, Access.READ, account=account, store_id=store_id,
                                 mode=mode, target=target, limit=limit, offset=offset)
        return format_response(data, "json")

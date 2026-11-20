"""MCP tool wrappers for folders."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import folders as folders_client
from outlook_mcp.client.mailbox import Access
from outlook_mcp.schemas import AccountSelector, StoreSelector
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_folders",
        annotations={
            "title": "List Outlook folders",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_folders(
        root: Annotated[
            Optional[str],
            Field(description="Folder to start from. Defaults to the selected store root."),
        ] = None,
        max_depth: Annotated[int, Field(ge=1, le=10)] = 4,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
        account: AccountSelector = None,
        store_id: StoreSelector = None,
    ) -> str:
        """List the folder tree with item and unread counts."""
        items = await bridge.call(
            folders_client.list_folders, Access.READ,
            account=account,
            store_id=store_id,
            root=root,
            max_depth=max_depth,
        )
        payload = {"count": len(items), "items": items}
        return format_response(payload, response_format)

    @mcp.tool(
        name="outlook_create_folder",
        annotations={
            "title": "Create Outlook folder",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_create_folder(
        name: Annotated[str, Field(min_length=1, description="Name of the new folder.")],
        parent: Annotated[str, Field(description="Parent folder. Defaults to inbox.")] = "inbox",
        account: AccountSelector = None,
        store_id: StoreSelector = None,
    ) -> str:
        """Create a sub-folder under the given parent."""
        data = await bridge.call(
            folders_client.create_folder, Access.WRITE,
            account=account,
            store_id=store_id,
            parent=parent,
            name=name,
        )
        return format_response(data, "json")

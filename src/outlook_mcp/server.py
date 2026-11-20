"""FastMCP server construction and tool registration."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.config import ToolPolicy, load_tool_policy
from outlook_mcp.tools import register_all
from outlook_mcp.ui import register_ui

INSTRUCTIONS = """\
This MCP server provides configurable access to Microsoft Outlook desktop on
Windows via COM automation. Depending on the local tool policy, it can read
and send mail, manage calendar events, contacts, tasks, categories and rules,
and check Out-of-Office status.

Authentication piggybacks on whatever account Outlook is signed into —
no Microsoft Graph API, no Entra app registration, no OAuth tokens.

PREREQUISITE: classic Outlook (OUTLOOK.EXE) must be installed. The
server attaches to Outlook at startup and can launch it; the new "modern" Outlook
(olk.exe) is NOT supported.

Only use tools advertised by this server; the local config may disable tools.
Use outlook_whoami to discover authorized accounts and their StoreIDs. For multiple
accounts specify account (exact SMTP address) or store_id on each mailbox call.
Return the store_id alongside the entry_id on follow-up calls.

Tool categories (all prefixed `outlook_`):
  - Mail: list, search, get, send, reply, forward, move, delete, mark, save_attachments
  - Folders: list_folders, create_folder
  - Calendar: list_events, get_event, create_event, update_event, delete_event, respond_event
  - Contacts: list_contacts, search_contacts (incl. org directory/GAL),
    get_contact, resolve_name (name -> SMTP address)
  - Tasks: list_tasks, create_task, complete_task
  - Categories: list_categories, set_category
  - Rules: list_rules, toggle_rule  (modifies live mail rules — confirm first)
  - Out-of-Office: get_out_of_office
  - Account: whoami  (sanity check on the bound mailbox)

Most read tools accept response_format='markdown' (default) or 'json'.
Item references use Outlook EntryID strings; list tools return them on
every item — pass them back to detail/edit/delete tools.

All datetimes are in the USER'S LOCAL timezone with an explicit UTC
offset (e.g. 2026-06-10T16:33:22+05:00). Present them as-is — do NOT
convert to another timezone. outlook_whoami reports the timezone name
and current local time.
"""


def build_server(*, policy: ToolPolicy | None = None) -> tuple[FastMCP, OutlookBridge]:
    """Construct the FastMCP instance, bridge, and wire all tools.

    The bridge is *not* started here — entrypoint.main() does that so
    the readiness wait happens after logging is configured.
    """
    # Validate config before constructing or starting any Outlook components.
    effective_policy = policy if policy is not None else load_tool_policy()
    mcp = FastMCP("outlook_mcp", instructions=INSTRUCTIONS)
    bridge = OutlookBridge()
    register_all(mcp, bridge, effective_policy)
    register_ui(mcp)
    return mcp, bridge

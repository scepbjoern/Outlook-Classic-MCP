# Local MCP tool configuration

By default, the server exposes all 31 Outlook tools. To restrict which MCP
operations are advertised and callable, create a **local user-specific JSON**
configuration file. No repository configuration is required.

## Config file location and startup

On Windows, the default is `%APPDATA%\outlook-classic-mcp\config.json`
(typically `C:\Users\<user>\AppData\Roaming\outlook-classic-mcp\config.json`).
On other platforms, it is `$XDG_CONFIG_HOME/outlook-classic-mcp/config.json`
or `~/.config/outlook-classic-mcp/config.json` when `XDG_CONFIG_HOME` is unset.
These paths refer to the machine **running the MCP server**, not the LLM host.

If the default file does not exist, the historic full tool set remains enabled.
To select a different file, set `OUTLOOK_MCP_CONFIG` **before** starting the
server. For example, in PowerShell:

```powershell
$env:OUTLOOK_MCP_CONFIG = 'C:\MCP-config\outlook-tools.json'
outlook-mcp
```

For MCP clients that launch the server, configure the variable in the client's
server environment. Restart the server process after editing the file. Reconnect
the client and request `tools/list` again to refresh its cached tool descriptions.
The running server keeps the policy it loaded at startup.

**Fail-closed startup:** An explicitly configured path that is missing,
unreadable or empty causes an error and process exit code 2. Any existing but
invalid default config also aborts startup **before Outlook is started**.
Unknown groups/tools/fields, duplicate JSON keys, unsupported versions, invalid
types, malformed JSON or invalid UTF-8 cause errors instead of being ignored.
Files are limited to 1 MiB.

## Format and precedence

Copy [the fictional example](../examples/config.example.json) to your local
config location and edit it. It disables only mail deletion and the Rules group:

```json
{
  "version": 1,
  "tools": {
    "disabled_tools": ["outlook_delete_mail"],
    "disabled_groups": ["rules"]
  }
}
```

`version` is optional and defaults to `1`. The optional `tools` object
supports four case-sensitive arrays:

| Property | Meaning |
| --- | --- |
| `enabled_groups` | If present, restrict to these groups (group allowlist). |
| `disabled_groups` | Deny every tool within these groups. |
| `enabled_tools` | If present, restrict to these individual tools (tool allowlist). |
| `disabled_tools` | Deny these individual tools. |

**Precedence:** Both allowlists must permit a tool when both are provided (the
intersection). Either deny list overrides every allowlist. An omitted allowlist
imposes no restriction, while an explicitly empty allowlist `[]` permits
nothing. An empty deny list denies nothing. Tool-specific allowances **cannot**
override a denied group. If a name occurs in both allow and deny lists, denial
wins.

For example, this limits access to **read-only calendar tools** and account
discovery:

```json
{
  "version": 1,
  "tools": {
    "enabled_groups": ["calendar", "account"],
    "disabled_tools": [
      "outlook_create_event", "outlook_update_event",
      "outlook_delete_event", "outlook_respond_event"
    ]
  }
}
```

Supported groups: `mail`, `folders`, `calendar`, `contacts`, `tasks`,
`categories`, `rules`, `out_of_office`, `account`. Prefix the names in
the [README Tools table](../README.md#tools) with `outlook_` when using them in a policy. The
authoritative tool inventory is `src/outlook_mcp/config.py`. If you add a tool
to the source code, add it to that inventory and update the registration tests;
an unlisted tool decorator otherwise rejects registration.

## Account selection and migration

Existing tool-only files use version 1, including files with no version field.
To add account permissions, set `version` to `2`, keep your `tools` settings
and add an `accounts` object. Both versions reject unknown properties. An
older server rejects version 2. On downgrade, restore the original tool-only
file and review the older server's mailbox access before starting it.

```json
{
  "version": 2,
  "tools": {"disabled_tools": ["outlook_delete_mail"]},
  "accounts": {
    "read": ["alpha@example.test", "beta@example.test"],
    "write": ["beta@example.test"]
  }
}
```

| Policy | One account loaded | Several accounts loaded |
| --- | --- | --- |
| Missing default file or version 1 | The sole account has read and write access. Selection is optional. | Mailbox operations require a version 2 policy. |
| Version 2 | The sole account is selected automatically. Its configured permissions apply. | Each call requires an explicit selector and the appropriate permissions. |

Version 2 requires `accounts`. Omitted or empty `read` and `write` lists grant
zero permissions for that access type. Each entry is an exact SMTP address.
The server validates the entire JSON file before starting COM. It checks
account membership, uniqueness and delivery-store identity against Outlook
during each call. An address absent from the loaded profile grants no access.

Call `outlook_whoami` for accessible accounts, their permissions and StoreIDs.
Pass `account` (SMTP address) or `store_id` (delivery StoreID). When supplied
together, both must identify the same unique account. Duplicate addresses
and shared delivery stores are ambiguous. Multiple accounts always require
a selector, even if only one is authorized. Use the `store_id` returned with an
`entry_id` for later operations. A move can produce a new EntryID.

Reading mailbox data, including saving attachments to disk, requires read
access. Creating mail or drafts, events, folders and tasks, updating event
fields and completing tasks require write access. Mail replies, forwarding,
moving, deletion and marking, meeting responses, event deletion, setting
categories and toggling rules require both rights because these actions also
return or transmit information about existing items.

## Calendar privacy

After upgrading and restarting, a missing config, version 1 (including an
omitted version), and existing version 2 files all default to `busy_only`.
The new setting governs results from native calendar appointments. Tool
registration and account read/write permissions still apply independently.

| Effective mode | Native Private (2) | Unreliable or unknown sensitivity |
| --- | --- | --- |
| `busy_only` (default) | “Appointment”, start, end and busy status only | Same |
| `exclude` | Omitted from calendar lists. Direct detail returns the same error as a missing item. | Same |
| `full` | Full fields, including subject, ID and body where provided by the tool | Availability only |

To restore the previous detail view for **known** private appointments in a
single-account installation, add `"calendar_privacy": {"default": "full"}`
to the existing file. To configure individual accounts, use version 2.
For example, this configuration grants both stores read access, permits edits
in one, and hides private appointments in the other.

```json
{
  "version": 2,
  "accounts": {
    "read": ["alpha@example.test", "beta@example.test"],
    "write": ["alpha@example.test"]
  },
  "calendar_privacy": {
    "default": "busy_only",
    "accounts": {"beta@example.test": "exclude"}
  }
}
```

The per-account key is the exact SMTP address of the authorized delivery
store, matched with the same case rules as account permissions. Selection
by StoreID uses the same policy. A global `exclude` takes precedence over
every override. Invalid settings abort startup. Older server versions reject
the new configuration field. Review the older version's disclosure behaviour
before downgrading.

The availability record has exactly four fields. `busy_status` uses Outlook's
codes 0 (free), 1 (tentative), 2 (busy), 3 (out of office) and 4 (working
elsewhere). Unreadable availability values are `null`. The list envelope
identifies the selected store and requested date range. Its count and limit
apply only to visible rows. Under `exclude`, raw folder item and unread totals
are also suppressed because mixed folders can contain hidden appointments.

Each expanded occurrence or exception uses its own sensitivity. When Outlook
resolves an EntryID to the series master, the result describes the master.
Its disclosure is limited by the strictest live exception. Inaccessible
recurrence information triggers the unknown-sensitivity fallback.
[Outlook recurrence exceptions](https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern.exceptions)

Create, edit, delete, meeting-response and category operations on protected
appointments return a status acknowledgement. A `busy_only` create also
returns the availability fields. Write permission still allows an explicit
change of sensitivity, including removing the Private flag. Subsequent reads
use the new classification. Assigning a color category named “Privat” has no
effect on classification. Separately stored meeting requests in mail folders
follow the mail tools' permissions.

## Outlook provider behaviour

- Each account maps to its delivery store in the loaded MAPI profile. Separate profiles require separate Outlook sessions. Archive PSTs and delegate-only stores without an account mapping cannot be selected.
- Calendar events, mail drafts and tasks are created in the selected store's corresponding folder. Sending binds the Outlook account. Use exact SMTP recipient addresses in multi-account profiles. Item lookup supplies the StoreID and verifies the returned parent folder before accessing the item.
- Folder paths and move destinations stay within the selected store. Prefixing a path with a store display name cannot select a different account.
- Saved contacts come from the selected store. Exchange directory search compares the provider section UID of the store with that of each Global Address List. An unavailable or ambiguous match leaves `searched_directory` false. Outlook's Ctrl+K name resolver is available only in a single-account, single-store profile. [Microsoft's store-specific address list lookup](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/address-book/identify-the-global-address-list-or-a-set-of-address-lists-with-a-store)
- Category listing uses the selected Exchange mailbox. Other providers require a single-account, single-store profile because secondary providers can return the primary category list. [Microsoft's Store.Categories reference](https://learn.microsoft.com/en-us/office/vba/api/outlook.store.categories)
- Rules use the selected store's rule collection. Enabling a rule in a multi-account or multi-store profile requires inspectable, mailbox-local actions. Move and copy destinations must belong to the same store. OOO reads the selected store's provider property and reports `unknown` if unavailable.
- If Outlook disconnects during a write, inspect the selected store before retrying. A provider may have completed part of the operation, such as saving a meeting before sending its invitation.

## Security boundary

Disallowed tools are filtered **at registration time** and are absent from MCP
`tools/list`. Direct calls to those names cannot invoke Outlook operations.
MCP App resources are static and may still display controls for disabled tools;
those controls cannot invoke unregistered tools.

The server checks account rights on every registered mailbox operation. To allow
calendar reads and event creation only, use a version 2 policy with
`tools.enabled_tools` set to `outlook_whoami`, `outlook_list_events`,
`outlook_get_event` and `outlook_create_event`. Give write access only to
the account intended for new events. Keep the account lists local.

Interactive clients can show the proposed event and ask for one-time approval.
In [Codex](https://learn.chatgpt.com/docs/extend/mcp), set
`mcp_servers.<id>.tools.outlook_create_event.approval_mode` to `"prompt"`
and use human review for the call. A client calling the server directly can
create events in a write-authorized account without that prompt. Enabling
`outlook_create_event` also permits invitations and recurring events. The
tool allowlist filters names only. Server-side approval for individual actions
is discussed in [issue #3](https://github.com/scepbjoern/Outlook-Classic-MCP/issues/3).

Allowed tool outputs may also be transmitted to the MCP client or its LLM
provider; running the COM bridge locally does not prevent that.

Whoever can edit, replace or remove the policy or change the server environment
can change the tool set at the next startup. For a mandatory policy, set
`OUTLOOK_MCP_CONFIG` explicitly so a missing file stops startup. Protect the
file, its parent directory and the client's launch settings with local filesystem
permissions. Store personal policies in an untracked location. Never commit
personal mailbox names, customer identities, credentials or secrets into an
example configuration.

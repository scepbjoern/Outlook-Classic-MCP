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

## Security boundary

Disallowed tools are filtered **at registration time** and are absent from MCP
`tools/list`. Direct calls to those names cannot invoke Outlook operations.
MCP App resources are static and may still display controls for disabled tools;
those controls cannot invoke unregistered tools.

**Hiding tools is not account authorization or user approval.** Every enabled
tool retains the Outlook COM permissions of the signed-in Windows user,
including access to multiple stores where the tool implementation allows it.
These rules do not enforce account allowlists, inspect recipients, require
write approval, or redact information returned by allowed tools. Those
protections are separate work in [issue #2](https://github.com/scepbjoern/Outlook-Classic-MCP/issues/2)
and [issue #3](https://github.com/scepbjoern/Outlook-Classic-MCP/issues/3).
Allowed tool outputs may also be transmitted to the MCP client or its LLM
provider; running the COM bridge locally does not prevent that.

Whoever can edit, replace or remove the policy or change the server environment
can change the tool set at the next startup. For a mandatory policy, set
`OUTLOOK_MCP_CONFIG` explicitly so a missing file stops startup. Protect the
file, its parent directory and the client's launch settings with local filesystem
permissions. Store personal policies in an untracked location. Never commit
personal mailbox names, customer identities, credentials or secrets into an
example configuration.

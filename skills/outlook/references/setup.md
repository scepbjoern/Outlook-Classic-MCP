# Connect a client to classic Outlook

The server runs in the Windows user session that has classic Outlook and the intended profile. A compatible MCP host starts the installed fork as a child process and exchanges MCP messages over standard input and output (STDIO). A client on another device can use this path only if its execution host actually runs the process in that Windows session.

Select the fork installation by its reviewed **full 40-character Git commit**. The public PyPI distribution `outlook-classic-mcp==0.3.3` belongs to upstream. Follow the installation instructions for the selected commit. [Fork installation procedure](../../../docs/installation.md)

Set the STDIO command to the absolute path of that installation's `.venv/Scripts/python.exe`. Pass `-m` and `outlook_mcp` as two separate arguments, and set `OUTLOOK_MCP_CONFIG` to the absolute path of the prepared local policy in that process's environment. The example paths below are fictional. Create the policy before connecting. An explicitly selected missing or invalid policy stops startup. A missing default policy can expose all tools.

| Client / execution environment | How to connect and where results go |
| --- | --- |
| OpenCode on Windows | Configure a local MCP server with an executable-and-arguments array. OpenCode can include results in requests to the selected model provider. [OpenCode MCP](https://opencode.ai/docs/mcp-servers/) |
| Codex or ChatGPT desktop with local MCP settings | Use the Windows execution environment and the STDIO command, arguments and environment above. Codex supports a per-tool human prompt. Returned data can reach OpenAI or the configured provider. [Local MCP settings and approvals](https://learn.chatgpt.com/docs/extend/mcp) |
| VS Code with a local MCP host | Use the user MCP configuration, a named entry under `servers`, `type: "stdio"`, `command`, `args` and `env`. In a remote workspace verify where the server runs. Copilot or another selected model can receive tool results. [VS Code MCP servers](https://code.visualstudio.com/docs/agent-customization/mcp-servers) |
| OpenClaw with a compatible Windows Gateway | In a Gateway version supporting local MCP servers, use `mcp.servers.outlook` with `transport: "stdio"`, `command`, `args` and `env`. The Gateway must be able to start the process in the Outlook user session. A Linux/WSL or remote Gateway alone cannot access Windows Outlook COM. The configured agent/model and any channel receiving the answer form part of the disclosure boundary. [OpenClaw MCP](https://docs.openclaw.ai/tools/mcp) · [Transport fields](https://docs.openclaw.ai/cli/mcp/transports) |
| Hosted ChatGPT / cloud-only session | A web chat does not start this process on the user's PC. ChatGPT custom MCP supports remote endpoints and, for eligible setups, Secure MCP Tunnel with a separately configured local client. That route carries requests and results through OpenAI. This project supplies a STDIO server. A tunnel or other remote deployment requires a separate access and privacy decision. [ChatGPT custom MCP](https://help.openai.com/en/articles/12584461-connecting-custom-mcp-servers-in-chatgpt) · [Secure MCP Tunnel](https://developers.openai.com/api/docs/guides/secure-mcp-tunnels) |

When using the bundled plugin, set `OUTLOOK_MCP_PYTHON` to the same installed interpreter. The client reads the policy from the server process environment.

### Example for OpenCode

In the user-local `opencode.json`, OpenCode V2 groups named servers under `mcp.servers`. The Windows connection was checked with OpenCode 2.2.7 and a synthetic Outlook bridge. OpenCode V1 expects the same named definition directly under `mcp`. [V1 and V2 configuration](https://opencode.ai/docs/migrate-v1/#mcp-servers)

```json
{
  "mcp": {
    "servers": {
      "outlook": {
        "type": "local",
        "command": ["C:/Tools/OutlookMCP/.venv/Scripts/python.exe", "-m", "outlook_mcp"],
        "environment": {"OUTLOOK_MCP_CONFIG": "C:/MCP-local/policy.json"}
      }
    }
  }
}
```

### Example for Codex

In the user's `config.toml`, use the installed fork and policy. The per-tool setting below requests interactive approval when the create tool is enabled. Select human review for that call.

```toml
[mcp_servers.outlook]
command = "C:/Tools/OutlookMCP/.venv/Scripts/python.exe"
args = ["-m", "outlook_mcp"]

[mcp_servers.outlook.env]
OUTLOOK_MCP_CONFIG = "C:/MCP-local/policy.json"

[mcp_servers.outlook.tools.outlook_create_event]
approval_mode = "prompt"
```

### Discover tools before reading Outlook

After upgrading for all-day support, inspect the published input schemas on a fresh connection. Both calendar write tools should advertise optional boolean `all_day`.

Have the client initialize the STDIO connection and inspect its current tool list (MCP `tools/list`). This is a protocol operation. Use the names and schemas offered by that connection. An absent tool can reflect server policy or a client filter. A skill file, config entry or static MCP App view alone establishes no permissions. A successful listing establishes tool discovery only. Use the enabled `outlook_whoami` tool to check local readiness and the selected account's Exchange connection. A connected account or local save does not establish that a particular item has synchronized. See [readiness and write receipts](../../../docs/outlook-readiness.md) and the [optional Windows check](../../../docs/outlook-readiness-smoke.md).

After a server upgrade or policy edit, end the old server process, reconnect and refresh the client's tool descriptions. For calendar availability, verify that the enabled create and update tools advertise optional integer `busy_status` with range 0–4 on the fresh connection. For a first read, follow the [calendar walkthrough](calendar-workflow.md). It includes a restrictive policy and the expected disclosures.

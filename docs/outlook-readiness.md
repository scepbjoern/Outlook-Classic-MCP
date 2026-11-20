# Outlook readiness, writes and Exchange synchronization

Start the MCP server and classic Outlook in the same interactive Windows user
session. The server starts its COM worker in the background. Discovery works
while Outlook opens. Use the enabled `outlook_whoami` tool as a read-only check.
Tool and account policy still apply during recovery.

| What the host reports | What it establishes | What to do |
| --- | --- | --- |
| Starting | The worker is opening Outlook and waiting for a usable MAPI session. The requested operation did not start. | Allow a cold launch to finish. Open classic Outlook manually if it stays closed. |
| Starting or reconnecting while a dialog is visible | Outlook may block inside a COM call. The server cannot identify which dialog is open. | Unlock the desktop, choose the intended profile yourself and complete any sign-in or security prompt. Check readiness again. |
| Unavailable | The attach polling period expired or Windows COM could not initialize. | Check Outlook in the same desktop session. A subsequent tool call can trigger a fresh attach on the worker. If COM initialization failed, repair the installation and restart the MCP host. |
| Ready | The local MAPI session responded. `accounts` lists only authorized delivery stores and their permissions. | Confirm the intended account and StoreID. Use them on mailbox calls. |
| Reconnecting | A provider error or stale-proxy probe detected loss of the session. New operations during reattachment do not start. | Wait for readiness and check the selectors. Work queued for an earlier session is rejected. |
| Read timed out | A read exceeded the 60-second call budget. A blocked COM call can outlive the caller. | Check the foreground Outlook window, then request the read again. |
| Timed out in the queue / operation did not start | The provider operation was never dispatched and cannot execute later. | Check readiness and submit the intended request again when ready. |
| Outcome is unknown | The operation began and the server lacks a complete acknowledgement. A timed-out COM call may still finish. | Allow the pending call to settle. Inspect the exact store, including Drafts, Calendar, Outbox and Sent Items as applicable. Check the output directory for attachment exports. Do not automatically repeat the action. |

An attach polls for up to 60 seconds between returning COM calls. Windows can
keep an individual call blocked in a dialog for longer. The server leaves that
worker in place. It does not start a second worker, choose a profile, answer a
prompt or terminate Outlook. A locked desktop, a service account or a scheduled
task without an interactive logon can leave prompts inaccessible. Use an
unlocked interactive session for this integration. Separate Outlook profiles
are separate sessions. To change profiles, settle pending writes, close Outlook,
select the desired profile manually and check the authorized stores again.

## Account diagnostics

`profile_name` is the current Outlook profile when the provider supplies a name
and all loaded accounts are authorized. Otherwise it is null. Each authorized
account has its own `exchange_connection`, derived from Outlook's read-only
Account property. One account's mode does not describe another store.

| Values | Interpretation |
| --- | --- |
| `offline`, `cached_offline` | Outlook reports offline mode for this account. Cached local changes can await synchronization. |
| `disconnected`, `cached_disconnected` | The Exchange connection is down. A local store may still answer. |
| `online`, `cached_connected_headers`, `cached_connected_drizzle`, `cached_connected_full` | The account reports an Exchange connection. Individual pending items may still be unsynchronized. |
| `not_exchange`, `unknown` | Outlook reports a non-Exchange account, or the property is unsupported, unavailable or unrecognized. |

Microsoft documents these properties in [Account.ExchangeConnectionMode](https://learn.microsoft.com/en-us/office/vba/api/outlook.account.exchangeconnectionmode), [OlExchangeConnectionMode](https://learn.microsoft.com/en-us/office/vba/api/outlook.olexchangeconnectionmode) and [NameSpace.CurrentProfileName](https://learn.microsoft.com/en-us/office/vba/api/outlook.namespace.currentprofilename).

## Write receipts and replay

On a successful mailbox write, `status` (where present) describes the existing
action result. `local_outcome` and `exchange_sync` describe the evidence.

| Local outcome | Evidence from Outlook |
| --- | --- |
| `saved` | The draft or appointment Save returned successfully. |
| `submitted` | Outlook's Send returned for mail, a forward, a reply or a meeting response. The item can still be in Outbox. |
| `saved_and_submitted` | Outlook saved the appointment and its invitation Send returned. |
| `completed` | The other requested Outlook operation returned, for example update, delete, move or folder creation. |

Every successful mailbox write returns `exchange_sync: "unverified"`, including
online writes. A local Save, Send, EntryID, local readback or starting Send/Receive
cannot establish per-item Exchange synchronization or recipient delivery. The
bridge neither triggers synchronization nor issues a verified Exchange receipt.
An independent server-backed observation can verify a specific change separately.
For example, see the [isolated Windows smoke test](outlook-readiness-smoke.md).

On disconnect, an eligible read can run once more within the original call
budget if the bridge confirms the same loaded profile. It resolves permissions
again using the account and store pinned when the original lookup succeeded.
Side effects, including exporting attachments to disk, never receive this
automatic replay. If a meeting Save returned before a later invitation step
failed, the error reports the known local save and the uncertain remaining
outcome. Keep the original appointment while checking the selected store.

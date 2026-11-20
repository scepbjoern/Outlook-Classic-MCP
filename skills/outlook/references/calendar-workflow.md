# Read a calendar, then preview one write

All accounts, paths, times and results here are fictional. Keep real policies and personal references outside the checkout. This walkthrough uses the installed fork's STDIO server. Outlook may sync independently. The client receives the server's actual results and can include them in model requests, conversation history or a channel response. Check the selected host, model provider and audience before requesting mailbox data.

## Read-only calendar

Save this version 2 server policy locally, for example at `C:/MCP-local/policy.json`, and reference it from the [client launch settings](setup.md). It allows identity and calendar reads for two synthetic mailboxes.

```json
{
  "version": 2,
  "tools": {"enabled_tools": ["outlook_whoami", "outlook_list_events", "outlook_get_event"]},
  "accounts": {"read": ["a@example.test", "b@example.test"], "write": []},
  "calendar_privacy": {"default": "busy_only", "accounts": {"b@example.test": "exclude"}}
}
```

1. Inspect `tools/list`. With this policy the server advertises exactly the three names above. A create button or an old client description grants no additional access.
2. Call `outlook_whoami` with `{"response_format":"json"}`. Suppose the result pairs `a@example.test` with `store-a`, and `b@example.test` with `store-b`. Both entries have `read: true` and `write: false`. Use the actual StoreID from the selected entry.
3. For “Show my availability in calendar A on 23 November”, call `outlook_list_events` with the arguments below, if the policy and the client's data flow permit a calendar list. Times without an offset refer to the Outlook host's local time. The interval ends at the following midnight.

```json
{"store_id":"store-a","start":"2026-11-23","end":"2026-11-24","limit":20,"response_format":"json"}
```

A native Private appointment in A returns four fields, for example:

```json
{"subject":"Appointment","start":"2026-11-23T10:00:00+01:00","end":"2026-11-23T11:00:00+01:00","busy_status":2}
```

Report the interval as busy. The list envelope still identifies the store and date range. There is no ID on this row for a detail lookup. A known Private item in B is omitted under `exclude` and a direct lookup looks missing. An empty list from B therefore cannot establish that the whole day is free. Respect the visible result and the requested limit.

**What can reach the model?** Under `busy_only`, only Private (2) and unreadable/unknown sensitivity are reduced to availability. Standalone items with Normal (0), Personal (1) or Confidential (3) can include subject, organizer, location and a body preview. A category named “Private” has no effect on this decision. With `full`, known Private items can expose details, while unknown sensitivity stays availability-only. Under `exclude` both Private and unknown are hidden. A global `exclude` takes precedence over account overrides.

“Summarize only free/busy” limits the final answer. If only availability may **leave the device**, a mixed-calendar list sent to a hosted model already exceeds that boundary. Use a verified local processing path that releases only permitted fields, or stop before the read. The server has no all-events availability projection. For a mail metadata request, `outlook_get_mail` with `include_body=false` and `include_html=false` omits the full body and HTML. The detail still contains identifying metadata and attachment names. Mail list and search tools can also return body previews. Check the schema and permitted payload first.

## Local aliases and categories

A user may keep `C:/MCP-local/outlook-rules.md` outside the checkout and explicitly make it available to their client. For example, a line could read “In b@example.test, the shorthand focus means the Outlook category Focus block.” Choose a location the user controls and edit that file when their labels change. The server reads its JSON policy only. Personal aliases are client guidance and do not belong in that JSON or the bundled skill. Loading a reference in a hosted model can disclose its contents.

If categories are requested, the server must advertise `outlook_list_categories` and the target account must permit reads. List categories in the **selected store**, verify the exact name and use that returned spelling. If missing or unverifiable, report it before proposing a write. For a standalone event, `add_categories: ["Focus block"]` preserves other assignments. For a series, category edits require explicit whole-series intent and `scope="series"`. Both category tools reject a single occurrence or exception, even when a time change is included in the same edit. Ask how to proceed without expanding the scope.

## One explicitly approved appointment

For a separate write-enabled session the operator can add **only** `outlook_create_event` to the example's tool allowlist and grant `accounts.write` to `b@example.test` only. Leave the read list intact. Restart, reconnect, repeat discovery and check the account rights. A server operator's grant enables the tool. The person using the client still reviews the proposed action.

> **Client:** Create “Focus block” in b@example.test's Calendar on 24 November 2026, 10:00–10:30 CET (UTC+01:00)? One appointment, native Private (2), 15-minute reminder, no attendees or invitations, no recurrence, category, location or notes.
>
> **User:** Yes, create that appointment as shown.

This is a sample interaction. Wait for the actual user's explicit reply to the preview or their one-time approval of the exact call. After a change, show the revised proposal. On refusal or silence, stop. Only after approval call `outlook_create_event` once with:

```json
{"store_id":"store-b","subject":"Focus block","start":"2026-11-24T10:00:00+01:00","end":"2026-11-24T10:30:00+01:00","sensitivity":2,"reminder_minutes":15}
```

B's `exclude` policy acknowledges this private appointment with `status="created"` and omits identifying event fields. Report the local creation from `local_outcome="saved"`. The receipt reports `exchange_sync="unverified"`. [Write receipt reference](../../../docs/outlook-readiness.md) A subsequent list omits it. If completion is uncertain, check the selected calendar in an authorized view before any retry. A direct MCP caller with write permission can bypass the client's human review. With attendees the same tool can send invitations, and with recurrence it can create a series. The create allowlist alone covers all three forms.

A one-off request ends with the reported result. Copying events between calendars on a schedule would require separate decisions about identity, privacy, conflicts, retries and authorization.

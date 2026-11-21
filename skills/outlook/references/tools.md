# Tool reference

All mailbox tools accept `account` (exact SMTP address) and `store_id`. Supply either selector in a multi-account profile. Use `outlook_whoami` for authorized selectors and keep each StoreID with its EntryID.

Use the active connection's `tools/list` for available names and input schemas. The entries below describe operations that may be enabled. Calendar return examples with identifiers apply only when disclosure permits them.

## Contents

- [Mail](#mail) — list_mails, search_mails, get_mail, draft_mail, send_mail, reply_mail, forward_mail, move_mail, delete_mail, mark_mail, save_attachments
- [Folders](#folders) — list_folders, create_folder
- [Calendar](#calendar) — list_events, get_event, create_event, update_event, delete_event, respond_event
- [Contacts](#contacts) — list_contacts, search_contacts, get_contact, resolve_name
- [Follow-ups](#follow-ups) — query_due
- [Tasks](#tasks) — list_tasks, create_task, complete_task
- [Categories](#categories) — list_categories, set_category
- [Rules](#rules) — list_rules, toggle_rule
- [Out-of-Office](#out-of-office) — get_out_of_office
- [Account](#account) — whoami
- [Common return-field glossary](#common-return-field-glossary)

---

## Mail

### `outlook_list_mails`

List mail items from a folder, newest first. Read-only.

| Param            | Type      | Default     | Notes |
| ---------------- | --------- | ----------- | ----- |
| `folder`         | string    | `"inbox"`   | Well-known name or path. Use a name or path within the selected store. |
| `limit`          | int 1–100 | `25`        | Max items to return. |
| `offset`         | int ≥0    | `0`         | Skip this many before returning. Use with the returned `next_offset` to paginate. |
| `unread_only`    | bool      | `false`     | If true, only unread mails. |
| `since`          | ISO-8601  | `null`      | Lower bound on `ReceivedTime`. |
| `until`          | ISO-8601  | `null`      | Upper bound on `ReceivedTime`. |
| `from_address`   | string    | `null`      | **Substring** match on sender. See gotcha re: `EX:/O=...` addresses. |
| `response_format`| `markdown`/`json` | `markdown` | Use `json` to extract `entry_id`s. |

**Returns** (`json` shape): `{ folder, count, offset, limit, items: [...], has_more, next_offset }`. Each item has: `entry_id, subject, from, from_address, to, received, unread, has_attachments, importance, preview` (200-char body excerpt).

### `outlook_search_mails`

Search a single folder by subject/body, subject-only, sender, or raw DASL. Read-only.

| Param            | Type     | Default          | Notes |
| ---------------- | -------- | ---------------- | ----- |
| `query`          | string   | required         | Search words (ALL must match, any order) — or a DASL @SQL filter when `scope='dasl'`. |
| `folder`         | string   | `"inbox"`        | Where to search. |
| `scope`          | enum     | `"subject_body"` | `subject_body` (default), `subject`, `from`, or `dasl`. |
| `limit`          | int 1–100| `25`             | |
| `response_format`| str      | `markdown`       | |

**Returns**: `{ query, scope, folder, count, items: [...] }`. Items have the same summary shape as `list_mails`.

Multi-word queries match items containing **all** the words (not the exact phrase), so `"teams not working"` finds "Synthetic report: Teams is not working". `scope='from'` matches display name, raw address, **and** the real SMTP address (works for Exchange senders too).

`scope='dasl'` is for power use — pass a complete `@SQL=...` filter and the server applies it raw. Only reach for this when subject_body/subject/from can't express what the user wants.

### `outlook_get_mail`

Fetch the body, all headers, and the attachment manifest for one mail. Read-only.

| Param            | Type   | Default | Notes |
| ---------------- | ------ | ------- | ----- |
| `entry_id`       | string | required | From a list/search result. |
| `include_body`   | bool   | `true`   | If false, omits `body`. Useful when you only need metadata. |
| `include_html`   | bool   | `false`  | Adds the raw `html_body`. Usually huge — leave off unless you specifically need the markup. |
| `max_body_chars` | int ≥0 | `10000`  | Body truncation cap; `0` = unlimited. |
| `response_format` | str | `markdown` | |

**Returns**: `{ entry_id, conversation_id, subject, from, from_address, to, cc, bcc, received, sent, unread, importance, categories, attachments: [{index, filename, size_bytes}], body }` plus `body_truncated`/`body_total_chars` when the cap was hit (re-call with a higher `max_body_chars` to read more) and `html_body` when `include_html=true`.

`attachments[].index` is **1-indexed**; pass it to `save_attachments` to save a single file.

### `outlook_draft_mail`

Save one new message in the selected account's Drafts folder. Uses the
composition fields in the table below, except `save_only`. The optional
`signature=true` appends the account's default Outlook signature. The
selected account requires write permission. The client previews the exact
message and obtains one-time approval to save. The local receipt contains
`status: "saved_to_drafts"`, `local_outcome: "saved"`, `entry_id`, `store_id`,
`subject` and `exchange_sync: "unverified"`.

### `outlook_send_mail`

Compose and send a new mail, or save it to Drafts. Has external side effect.

| Param          | Type      | Default   | Notes |
| -------------- | --------- | --------- | ----- |
| `to`           | list[str] | required  | One or more SMTP addresses. |
| `subject`      | string    | required  | |
| `body`         | string    | required  | Plain text unless `html=true`. |
| `cc`           | list[str] | `null`    | |
| `bcc`          | list[str] | `null`    | |
| `html`         | bool      | `false`   | When true, `body` is HTML. |
| `attachments`  | list[str] | `null`    | Absolute paths under user profile. |
| `importance`   | enum      | `"normal"`| `low` / `normal` / `high`. |
| `save_only`    | bool      | `false`   | **Save to Drafts instead of sending.** |

**Returns** (sent): `{ status: "sent", to, cc, bcc, subject }`. (Drafts): `{ status: "saved_to_drafts", entry_id, subject }`.

Always confirm the recipient list and subject with the user before calling this tool unless they have explicitly authorized you to send.

### `outlook_reply_mail`

Reply (or reply-all) to an existing mail. The original message is appended below your body, the same way Outlook's Reply button does it. Has external side effect.

| Param         | Type      | Default | Notes |
| ------------- | --------- | ------- | ----- |
| `entry_id`    | string    | required | The mail being replied to. |
| `body`        | string    | required | Your reply text. The quoted original is appended automatically. |
| `reply_all`   | bool      | `false`  | If true, includes the original CC list. |
| `html`        | bool      | `false`  | |
| `attachments` | list[str] | `null`   | |

**Returns**: `{ status: "sent", reply_all, in_reply_to, subject }`.

This sends immediately. To stage a new message for review, verify the intended recipients and call `outlook_draft_mail` after confirmation.

### `outlook_forward_mail`

Forward an existing mail to new recipients with an optional note above. Has external side effect.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `entry_id` | string    | required | |
| `to`       | list[str] | required | |
| `body`     | string    | `""`     | Optional note prepended to the forwarded content. |
| `cc`       | list[str] | `null`   | |
| `html`     | bool      | `false`  | |

**Returns**: `{ status: "sent", forwarded, to, subject }`.

### `outlook_move_mail`

Move a mail to another folder.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `target_folder` | string | required | Well-known name or path. |

**Returns**: `{ status: "moved", new_entry_id, folder }`.

The `entry_id` changes when an item moves stores. **Use the returned `new_entry_id`** if you need to act on the moved item again.

### `outlook_delete_mail`

Soft-delete (moves to Deleted Items). Reversible by the user from Outlook.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_mark_mail`

Toggle read state and/or follow-up flag.

| Param      | Type | Default | Notes |
| ---------- | ---- | ------- | ----- |
| `entry_id` | string | required | |
| `read`     | bool/null | `null` | `true` = mark read, `false` = mark unread, `null` = no change. |
| `flagged`  | bool/null | `null` | `true` = flag for follow-up, `false` = clear flag, `null` = no change. |

**Returns**: `{ status: "updated", entry_id, unread, flagged }`.

### `outlook_save_attachments`

Save one or all attachments from a mail to a local directory.

| Param              | Type    | Default | Notes |
| ------------------ | ------- | ------- | ----- |
| `entry_id`         | string  | required | |
| `output_dir`       | string  | required | Absolute path under user profile. Created if missing. |
| `attachment_index` | int ≥1  | `null`  | 1-indexed. Omit to save all. |

**Returns**: `{ status: "saved", count, files: [absolute paths], output_dir }`.

---

## Folders

### `outlook_list_folders`

Walk the folder tree under a root.

| Param            | Type    | Default | Notes |
| ---------------- | ------- | ------- | ----- |
| `root`           | string  | `null`  | Folder to start from. Default = the selected mailbox root. |
| `max_depth`      | int 1–10| `4`     | How deep to walk. |
| `response_format`| str     | `markdown` | |

**Returns**: `{ count, items: [{name, path, item_count, unread_count, default_item_type}, ...] }`. The `path` strings are exactly what you pass back as a `folder` parameter elsewhere.

### `outlook_create_folder`

Create a sub-folder under a parent.

| Param    | Type   | Default   | Notes |
| -------- | ------ | --------- | ----- |
| `name`   | string | required  | New folder name. |
| `parent` | string | `"inbox"` | Where to put it. |

**Returns**: `{ name, path, entry_id }`.

---

## Calendar

### `outlook_list_events`

List calendar events in a date range, including expanded recurring instances.

| Param                 | Type    | Default        | Notes |
| --------------------- | ------- | -------------- | ----- |
| `start`               | ISO-8601| now            | |
| `end`                 | ISO-8601| `start + 14d`  | |
| `limit`               | int 1–200| `50`          | |
| `include_recurrences` | bool    | `true`         | If false, only the master entries — usually you want true. |
| `response_format`     | str     | `markdown`     | |

**Returns**: `{ start, end, count, store_id, items: [...] }`. Items include `entry_id, store_id, subject, start, end, location, organizer, is_recurring, all_day, preview` (200-char body excerpt), plus `sensitivity` and `sensitivity_label` from the listed appointment or occurrence.

The native Outlook levels are Normal (0), Personal (1), Private (2) and Confidential (3).
These fields describe full rows. By default a Private or unclassified row contains only `subject: "Appointment"`, `start`, `end` and `busy_status`. Under `exclude`, it is omitted. With `full`, confirmed Private appointments can include identifying fields. A custom color category named “Privat” is independent of the native marker.

### `outlook_get_event`

Full event detail, including attendees and their RSVP status.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |
| `response_format` | str | `markdown` | |

**Returns**: summary fields + `body, attendees: [{name, address, type, response}], reminder_minutes, categories`.

For get, update, delete, respond and set-category, pass the listed `store_id` and `entry_id`.
For operations supported on a recurring instance, also pass `scope="occurrence"`
and its `original_start` with offset. Category writes require whole-series scope.
This selects the original slot even after a move. The full row marks a modified occurrence
with `is_exception=true`. Use `scope="series"` alone for the master. A legacy detail
read without scope describes the object opened by ID. Recurring writes require explicit scope.
Detail uses the same privacy policy as the list. A returned master also respects its live exceptions. A hidden item produces the standard missing-item error.

### `outlook_create_event`

Create a calendar event or meeting invite. **Adding any attendee turns this into a meeting that is sent immediately on success — there is no draft state for meeting invites.**

| Param               | Type           | Default | Notes |
| ------------------- | -------------- | ------- | ----- |
| `subject`           | string         | required | |
| `start`             | ISO-8601       | required | |
| `end`               | ISO-8601       | required | |
| `location`          | string         | `null`  | |
| `body`              | string         | `null`  | |
| `attendees`         | list[str]      | `null`  | Email addresses. Adding any value here makes this a meeting and sends invites. |
| `is_online_meeting` | bool           | `false` | Reserved. This server does not create an online meeting link. |
| `reminder_minutes`  | int 0–10080    | `15`    | Minutes before start. |
| `recurrence`        | Recurrence obj | `null`  | Use the recurrence schema advertised by this connection. |
| `sensitivity`       | int 0–3        | omitted | Set the native Outlook level. Omit to use Outlook's default. |
| `categories`        | list[str]      | omitted | Initial assignments. Supply exact category names from the selected store. |

**Returns**: `{ status: "created", entry_id, store_id, subject, start, end }` for full disclosure.

Protected appointment actions return status only. A `busy_only` create adds the four availability fields. Other protected appointment writes return only status.

Confirm attendee list, times, recurrence and any explicitly selected sensitivity with the user before calling.

### `outlook_update_event`

Update the supplied fields on the selected target. Edit cadence in Outlook.
Changing series start or end requires a verified empty exceptions collection.

| Param      | Type     | Default | Notes |
| ---------- | -------- | ------- | ----- |
| `entry_id` | string   | required | |
| `subject`  | string   | `null`  | |
| `start`    | ISO-8601 | `null`  | |
| `end`      | ISO-8601 | `null`  | |
| `location` | string   | `null`  | |
| `body`     | string   | `null`  | |
| `sensitivity` | int 0–3 | omitted | Omit to preserve the current level. An explicit 0 sets Normal. |
| `add_categories` | list[str] | omitted | Append missing names. Repeating a name has no further effect. |
| `remove_categories` | list[str] | omitted | Remove the named assignments, preserving other labels, including ones retired from the store list. |
| `replace_categories` | list[str] | omitted | Set the complete list. An empty array clears all assignments. |

Omitting the category fields preserves the item's assignments. Add and remove can
be combined if their names are disjoint. Replacement must be supplied alone.
On create and update, copy names exactly from `outlook_list_categories` with
the same store selector. The server rejects unknown names (including removal
requests) and conflicting operations before editing other fields, saving an
appointment or sending an invitation. A color category named “Privat” leaves the
native sensitivity and server disclosure policy intact.

Pass each name as its own array element. Accents, Unicode and punctuation are
preserved. Outlook's appointment setter uses the Windows regional list
separator. If a name in the resulting assignment contains that separator, the
server rejects the change to avoid splitting a label. For recurring appointments,
category edits require explicit whole-series scope. A category request for a
single occurrence or exception rejects the entire edit before any fields change.
Keep the requested scope when reporting this error to the user.
[Microsoft's AppointmentItem.Categories reference](https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.categories)

**Returns**: `{ status: "updated", entry_id }`. If the event has attendees, Outlook may send an updated-meeting notification when this saves.

### `outlook_delete_event`

Delete the selected appointment, occurrence or entire series. Preview the affected attendees
and scope. The tool calls Outlook's native Delete without an additional send.

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_respond_event`

Respond to a meeting invite.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `response`      | enum   | required | `accept` / `tentative` / `decline`. |
| `send_response` | bool   | `true`   | Set false to record locally without emailing the organizer. |

**Returns**: `{ status: "responded", response }`.

---

## Contacts

### `outlook_list_contacts`

List saved contacts from **contact folders in the selected store**, sorted by full name within each folder.

| Param            | Type     | Default | Notes |
| ---------------- | -------- | ------- | ----- |
| `limit`          | int 1–200| `50`    | |
| `offset`         | int ≥0   | `0`     | |
| `response_format`| str      | `markdown` | |

**Returns**: `{ count, offset, items: [...], has_more }`. Items: `entry_id, full_name, email, company, job_title, mobile, folder`.

On corporate accounts the personal contact folders are often nearly empty — colleagues live in the **directory** (Global Address List), which this tool does not list. Use `search_contacts` to find people.

### `outlook_search_contacts`

Word search across saved contacts (name, email, company, job title) **and the org directory (GAL)**. All query words must match.

| Param               | Type    | Default | Notes |
| ------------------- | ------- | ------- | ----- |
| `query`             | string  | required | e.g. `"alex reed"` matches "Alex Morgan Reed". |
| `limit`             | int 1–100| `25`   | |
| `include_directory` | bool    | `true`  | Also scan the Exchange Global Address List. A few seconds on large directories. |
| `response_format`   | str     | `markdown` | |

**Returns**: `{ query, count, items, searched_directory }`. Each item has `source: "contacts" | "directory"`. Directory items carry `full_name, email (SMTP), company, job_title, mobile` but **no `entry_id`** — they aren't Outlook items, so don't pass them to `get_contact`.

### `outlook_get_contact`

Full contact record (saved contacts only — needs an `entry_id`). Returns the summary fields plus `business_phone, home_phone, address, notes`.

### `outlook_resolve_name`

In a single-account, single-store profile, resolve a display name, alias, or address to its SMTP address — same mechanism as typing a name in To: and pressing Ctrl+K. Use before sending when you only know a person's name.

| Param  | Type   | Default | Notes |
| ------ | ------ | ------- | ----- |
| `name` | string | required | Full names resolve best; short fragments are often ambiguous. |

**Returns**: `{ resolved: true, query, display_name, smtp_address }` or `{ resolved: false, query, note }`. Ambiguous names do **not** resolve — fall back to `search_contacts` to browse candidates.

---

## Follow-ups

### `outlook_query_due`

Read outstanding dated items. Supply an exact `account` or `store_id`, even for a single mailbox. For example, `{"store_id":"store-a","mode":"on_date","local_date":"2026-11-23"}`. Use `today` (default) or `overdue` without a date. The Outlook host's local calendar day starts at midnight. Overdue means due before today. The entire named day is inclusive, including on a daylight-saving transition.

The response is JSON. Each row contains `item_type` (mail, task or calendar), permitted `subject`, local `due_date` as YYYY-MM-DD and `status`. Mail and task rows include `entry_id`. Calendar reminder rows include `reminder_at` and `appointment_start` when available. Keep the enclosing `store_id` with each row. Results are ordered by date and EntryID. `limit` defaults to 50 (1–200). Pass `next_offset` as `offset` to continue until it is null. Each call reads the current views, so changes between pages can affect the order.

Read `coverage` alongside `items`. The `to_do` and `default_tasks` statuses apply to those local provider views. `complete` with zero matches means no verified outstanding dated items in that view. `partial` means some reads or store checks failed, possibly including the entire folder. Rows from successful reads remain usable. If To-Do is unavailable, mail coverage is incomplete. Calendar coverage applies to permitted current reminders in this store. For calendar rows the day is the local date of `NextReminderDate`, including after snoozing. Use `mode="active_reminders"` without a date to list all visible alerts across dates. Mail and task views then report `not_requested`. Snoozed and scheduled alerts remain available in the date queries. A closed alert drops out of the collection. Absence alone cannot prove a dismissal. [Outlook properties and provider scope](../../../docs/tool-configuration.md#due-query-coverage)

## Tasks

### `outlook_list_tasks`

List tasks from the default Tasks folder, sorted by due date.

| Param                | Type     | Default | Notes |
| -------------------- | -------- | ------- | ----- |
| `limit`              | int 1–200| `50`    | |
| `include_completed`  | bool     | `false` | Default hides done tasks. |
| `response_format`    | str      | `markdown` | |

**Items**: `entry_id, subject, due_date, start_date, complete, percent_complete, importance, status`.

### `outlook_create_task`

| Param        | Type     | Default   | Notes |
| ------------ | -------- | --------- | ----- |
| `subject`    | string   | required  | |
| `due_date`   | ISO-8601 | `null`    | |
| `body`       | string   | `null`    | |
| `importance` | enum     | `"normal"`| low/normal/high. |
| `reminder`   | ISO-8601 | `null`    | Sets a reminder time. |

**Returns**: `{ status: "created", entry_id, subject }`.

### `outlook_complete_task`

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

Marks the task 100% complete. Returns `{ status: "completed", entry_id }`.

---

## Categories

### `outlook_list_categories`

Returns the selected delivery store's color categories: `{ count, items: [{name, color}, ...] }`. Secondary non-Exchange providers can expose the primary list, so the server restricts these providers to single-account, single-store profiles. Calendar category validation uses the same provider checks.

### `outlook_set_category`

Replace the categories on mail, tasks, standalone appointments or an explicitly
selected whole series. Individual occurrences and exceptions reject category edits.

| Param        | Type   | Default | Notes |
| ------------ | ------ | ------- | ----- |
| `entry_id`   | string | required | |
| `categories` | string | required | **Comma-separated names.** Empty string clears all. e.g. `"Important"` or `"Work, Follow-up"`. |

This **replaces** existing categories rather than adding to them. To add `Foo` to an item that already has `Bar`, send `"Bar, Foo"`. Get the current value first via `get_mail` / `get_event` if needed.

This legacy tool writes the supplied string directly, including unlisted names.
For appointments, `outlook_update_event` supports validated, store-specific
array parameters and can add or remove a label in one call.

---

## Rules

### `outlook_list_rules`

Returns the user's mail rules with their on/off state: `{ count, items: [{index, name, enabled}] }`.

### `outlook_toggle_rule`

| Param       | Type   | Default | Notes |
| ----------- | ------ | ------- | ----- |
| `rule_name` | string | required | **Exact** name from `list_rules`. |
| `enabled`   | bool   | required | `true` to enable, `false` to disable. |

This change is live the moment it's saved — no staging buffer. Confirm the rule name with the user before calling.

---

## Out-of-Office

### `outlook_get_out_of_office`

Reports whether OOO auto-reply is currently on. Returns `{ out_of_office: bool, status: "on"|"off" }`, or `{ out_of_office: null, status: "unknown", note: ... }` if the property isn't readable on this profile.

There is **no tool to enable, disable, or schedule OOO**. Tell users to manage it via Outlook → File → Automatic Replies.

---

## Account

### `outlook_whoami`

Returns `requires_account_selection`, authorized `accounts` (including `smtp_address`, `store_id`, `read`, `write`), `current_user`, `local_time`, `timezone` and `utc_offset`. With multiple accounts, choose the intended authorized store explicitly.

---

## Common return-field glossary

- `entry_id` — opaque, stable handle for an item. Pass back verbatim. Becomes invalid on delete; changes on cross-store move.
- `conversation_id` — groups mails in a thread. Same value across replies/forwards in one conversation.
- `from` — display name of the sender.
- `from_address` — sender address. **For Exchange senders, this is an `EX:/O=...` distinguished name, not SMTP.** Match by substring.
- Except for date-only `outlook_query_due` rows, `received` / `sent` / `start` / `end` / `due_date` — ISO-8601 strings **in the user's local timezone with explicit offset** (e.g. `2026-06-10T16:33:22+05:00`). Present as-is; never convert to another timezone.
- `unread` — bool. Note `mark_mail` returns `unread` (not `read`).
- `importance` — integer (0=low, 1=normal, 2=high).
- `categories` — comma-separated string of category names; empty string = none.
- `preview` — 200-char body excerpt. Not a substitute for `get_mail` / `get_event` when you need the full body.
- `has_more` / `next_offset` — pagination signals on list endpoints.

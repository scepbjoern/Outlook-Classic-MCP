# Gotchas

Every quirk and failure mode worth knowing, organized by *what the user sees*. Read this before doing anything that has external side effects, and consult the matching entry the moment something looks weird.

## "You said you sent it but I don't see it in Sent Items."

Check the selected account and the reported result. A missing message alone cannot establish whether Outlook saved or sent it. Check the authorized store before any retry and follow the operator's managed Outlook settings.

## Sender filter "from_address=alice@example.test" returns nothing

Exchange senders' `from_address` is **not an SMTP address**. It's an `EX:/O=ExchangeLabs/OU=...` distinguished name. The substring filter on `list_mails` matches against this raw string, which won't contain the SMTP address.

**Fix**: use `outlook_search_mails(scope="from", query="sarah@example.test")` — since v0.2.5 the `from` scope also matches the real SMTP address (via the MAPI sender-SMTP property), so it works for Exchange senders. For `list_mails`'s `from_address` filter, pass a name fragment that *will* appear in the DN — usually the first or last name.

## Times look wrong / shifted by a few hours

All datetimes the integration returns are in the **user's local timezone with an explicit UTC offset** (e.g. `2026-06-10T16:33:22+05:00`). Present them as-is — do **not** convert them to UTC or any other timezone. `outlook_whoami` reports the user's timezone name, current local time, and offset if you need to reason about it.

## "Search doesn't find the email I described"

`outlook_search_mails` matches items containing **all** the query words in any order — not the exact phrase, and not fuzzy. If a multi-word query misses, drop to the one or two most distinctive words. Keep in mind it searches **one folder at a time** (default: inbox) — check Sent or other folders explicitly if relevant.

## Toggling a rule changed the user's actual mail flow before they confirmed

`outlook_toggle_rule` modifies a live mail rule the moment it returns successfully. There's no staging buffer, no preview, no undo. If you call it on the wrong rule name, the user's mail is now being filed (or not filed) differently for real.

**Always**: `outlook_list_rules` first → confirm the **exact** rule name with the user → only then `outlook_toggle_rule`.

## "Set my OOO to on for the next two weeks" — and it just doesn't

The integration has `outlook_get_out_of_office` (read state) but no tool to set it. The auto-reply state lives in an Exchange property that requires EWS or Graph to write, neither of which the integration uses.

**Tell the user** to set OOO via Outlook → File → Automatic Replies. After they do, you can confirm it took effect with `outlook_get_out_of_office`.

## "Why did everyone get an email when I just rescheduled the meeting?"

Update and delete call Outlook's native Save and Delete on the selected appointment. Preview the affected meeting and scope. Delivery of updates or cancellations depends on Outlook and the account. These tools do not explicitly send notices.

**Workaround**: warn the user up front. If they want a quiet local-only change, they have to do it in Outlook directly with "Don't send updates" — the integration can't suppress it.

## Created a meeting and the invite went out before the user reviewed

`outlook_create_event` with `attendees=[...]` sends the invite **the moment the call succeeds**. There is no draft state for meeting invites the way there is for plain mail (`save_only=true`).

**Mitigation**: confirm subject, time, and attendee list with the user before calling. If the user wants to compose carefully, create the event without `attendees` first, then have them add attendees manually in Outlook.

## "I replied — can you save that to drafts so I can edit it first?"

`outlook_reply_mail` doesn't have a `save_only` flag. It sends.

**Workaround**: build the reply yourself and save it via `outlook_send_mail(save_only=true)`:

```
to = [verified sender SMTP address]                        # or all of original.to/cc for reply-all
subject = "Re: " + original.subject                 # don't double-prefix
body = your_reply + "\n\n" + original.body          # quote manually
outlook_send_mail(store_id=store, to=to, subject=subject, body=body, save_only=true)
```

Tell the user the draft is in the Drafts folder for them to review.

## Item not found / EntryID errors after a move

Moving an item with `outlook_move_mail` returns a `new_entry_id`. The original `entry_id` no longer resolves.

**Always** use `new_entry_id` for any subsequent action on the moved item. If you've lost track, re-list the destination folder.

## "Schedule it daily for the next 6 months" — and then they want to change cadence

`outlook_update_event` cannot modify the recurrence pattern. It supports field and category edits within the selected scope.

Edit the recurrence pattern in Outlook when cadence, interval or end date must change. The series may already contain moved, private or deleted occurrences.

## Set categories overwrote the existing tags

`outlook_set_category` **replaces** the categories field, it doesn't append. Categories are stored as a comma-separated string and the tool overwrites that whole string.

**Pattern**: read first, merge, then set.

```
existing = outlook_get_mail(store_id=store, entry_id=entry_id, include_body=false).categories
new_value = ", ".join(filter(None, [existing, "Important"]))
outlook_set_category(store_id=store, entry_id=entry_id, categories=new_value)
```

To **clear** all categories, pass an empty string.

## "Find Sarah's email address" returns nothing

On corporate accounts the personal Contacts folder is usually **empty** — colleagues live in the org directory (Global Address List). `outlook_search_contacts` searches both saved contacts and the directory (`include_directory=true` by default; directory hits have `source: "directory"` and no `entry_id`). For a direct name→address lookup, `outlook_resolve_name` is faster and authoritative; if it reports not-resolved the name was ambiguous — try the full name or fall back to `search_contacts`.

## "It says my category 'Urgent' doesn't exist."

`outlook_set_category` doesn't validate the names against the profile's defined categories. If you set a category that isn't defined, Outlook will accept the string but the item won't get the color and the user won't see a recognizable tag in the UI.

**Best practice**: `outlook_list_categories` first; only assign names that come back from there. If the user wants a brand-new category, they need to create it in Outlook (Home → Categorize → All Categories → New) — there is no `create_category` tool.

## Attachment / output_dir failures

The integration enforces a **user-profile sandbox** on all filesystem paths. Passing `D:\stuff\file.pdf` or `\\fileserver\share\foo` will fail with a sandbox error.

**Tell the user** their options:

1. Move the file under their user profile (`C:\Users\<them>\...`) — easiest.
2. Set the env var `OUTLOOK_MCP_ALLOW_ANY_PATH=1` in the integration's environment to disable the sandbox — they have to restart the integration after changing it. Use only when they really need it (e.g. corporate file shares).

## Path errors specifically about "must be absolute"

Relative paths like `Documents/file.pdf` always fail. Always pass `C:\\Users\\<user>\\Documents\\file.pdf` (or with forward slashes — both work).

## A write was interrupted

Inspect the selected store before repeating a call. A provider may have saved an item before the client received its result. Tool discovery alone cannot establish Outlook connectivity or sync status.

## DASL filter syntax errors when using `scope='dasl'`

`scope='dasl'` passes the entire `query` string into Outlook's `Items.Restrict` as a raw `@SQL=` filter — it must be valid DASL.

**Examples that work**:

```
@SQL="urn:schemas:httpmail:subject" LIKE '%budget%'
@SQL="urn:schemas:httpmail:hasattachment" = 1
@SQL="urn:schemas:httpmail:read" = 0 AND "urn:schemas:httpmail:importance" = 2
```

If you don't actually need DASL, use `scope='subject_body'` / `'subject'` / `'from'` instead — they construct the filter for you and handle escaping.

## Unread count looks wrong

`unread` on a mail item reflects whether the item is unread. `unread_count` on a folder (from `list_folders`) is Outlook's cached count and can lag the actual state by seconds, especially right after a sync.

If the numbers disagree, trust the per-item `unread` flag.

## Selecting an inbox in a multi-account profile

Use `outlook_whoami` to identify the authorized delivery store. Pass its SMTP `account` or `store_id` with `folder="inbox"`. Folder paths are relative to that store.

## Send error mentions "ResolveAll" or "recipient could not be resolved"

The address you passed to `attendees=` (or in `to`/`cc`/`bcc`) couldn't be resolved against the user's address book or directory. Possible causes:

- A typo in the address.
- An internal alias that the user's directory recognizes (`alice`) but Outlook needs the full SMTP form for (`alice@example.test`).
- An external address being rejected because the user's tenant blocks external invites.

**Recover**: ask the user for the full SMTP address; retry.

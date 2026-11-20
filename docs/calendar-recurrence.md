# Selecting a recurring appointment

A range read includes appointments overlapping the interval from the inclusive start
to the exclusive end. The default window is 14 days. Outlook sorts by Start, expands
recurrences and restricts the collection before iteration. The server checks exact
times, including offsets at daylight-saving transitions. An open-ended series needs
no end date. An excessive scan returns an error asking for a narrower window.
[Outlook collection semantics](https://learn.microsoft.com/en-us/office/vba/api/outlook.items.includerecurrences)

| Client intent | Fields to pass to get, update, delete, respond or set-category |
| --- | --- |
| Single appointment | Listed `store_id` and `entry_id` |
| One occurrence, including a modified exception | Same IDs, `scope="occurrence"` and the listed `original_start` |
| Entire series | Series IDs and `scope="series"`. Omit `original_start`. |

For example, a fictional occurrence originally scheduled at 09:00 CET on
23 November can move to 14:00 and keep the following detail arguments.

```json
{
  "store_id": "synthetic-store",
  "entry_id": "synthetic-series",
  "scope": "occurrence",
  "original_start": "2026-11-23T08:00:00+00:00"
}
```

The account policy binds each request to its delivery store. An exact SMTP account
and a StoreID supplied together must agree. Within that store, multiple occurrences
can share the master EntryID. For a regular instance the selector uses StartUTC.
For a live exception it uses the replaced slot in UTC. The server checks native
recurrence state and verifies the resolved item, store and original instant before
changing fields. Deleted slots and unverifiable instance lookups fail. If the provider
cannot supply an original instant, the listed row has no actionable ID. Relist after
a move between stores or a change of pattern. [Outlook GetOccurrence](https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern.getoccurrence),
[exception replacement time](https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/pidlidexceptionreplacetime-canonical-property)

Before a write, preview the account, the single occurrence or the entire series,
the permitted original and current times, the proposed change and the impact on
attendees. Obtain approval for that selection in the interactive client. A direct
client with account write permission can invoke the tool without this interaction.
Update and delete use Outlook's Save and Delete. The tools make no additional send
call for a meeting. A series time change requires a verified empty collection of
exceptions, including deleted ones. Change recurrence cadence in Outlook.

A legacy ID-only detail lookup still describes the object Outlook opens. A master
uses the strictest applicable exception privacy policy. Scoped detail reads classify
the selected instance. Under `busy_only`, private rows contain only availability.
Under `exclude`, private rows are absent and a known target reads as missing.

## Manual smoke test for classic Outlook and Exchange

Pending separate agreement for a designated test account. Use fictional appointments.
Any verification involving attendees requires separately agreed test mailboxes and
approval for the specific meeting action. Record outcomes without account or item IDs.

| Fixture and operation | Expected observation |
| --- | --- |
| Finite weekly series, first and last slot. Inspect range, detail and an adjacent out-of-pattern slot. | Each returned selector opens its own slot. Outside the pattern the lookup fails. |
| Open-ended daily series over a short window. | Results stay within the range and respect the visible limit. |
| Timed series across both clock changes, including a repeated hour, and a series in another time zone. | Local offsets and original UTC instants round-trip. A mismatched instant cannot select a neighbour. |
| All-day series across a clock change. Query inside the day and at the exclusive end. | The overlapping day is present. Local midnight boundaries survive a 23- or 25-hour day. |
| Move one instance, change its subject and mark it Private. Reopen by original slot in all three privacy modes. | Full detail follows the moved item. Restrictive modes yield availability or the missing-item error. The UI opens the selected row. |
| Delete one occurrence and reuse its former selector. Try a series time change after the deletion. | The slot is absent, neighbours remain, and both the stale write and series time change fail. |
| Try a bare-ID update/delete, conflicting selectors and a read-only store. Then update one occurrence and explicitly delete a separate test series. | Rejected calls leave items untouched. Each approved write affects exactly the previewed scope. |
| Disposable organizer and attendee copies, after specific agreement. | Check native Outlook results and delivery for one occurrence and for a series. Record which recipients saw an update or cancellation. |

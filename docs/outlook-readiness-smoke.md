# Optional Windows smoke test for Outlook readiness

Use a disposable interactive Windows VM with classic Outlook, two synthetic
Exchange accounts, a synthetic recipient you control and a separate browser
session with web access to the test mailbox. Configure one writable and one
read-only account and enable only the tools needed. Use a unique invented
subject per operation. Keep identifiers, profile names, mail and logs local.

1. **Cold launch and profile selection.** Close Outlook normally. Configure the
   test installation to ask for a profile, then start the MCP host. List tools
   and call `outlook_whoami` while the picker is open. Expect a starting
   message, no queued write and a responsive host. Choose the intended profile
   manually. Repeat the read until it reports ready and the expected authorized
   stores. If a sign-in dialog appears, complete it manually.
2. **Offline local save.** Put the synthetic Exchange account in Work Offline.
   Inspect its connection mode with `outlook_whoami`. Save one new draft with
   `outlook_draft_mail` and the test StoreID after client confirmation. Check for exactly
   one matching local draft and record the saved/unverified receipt. Check
   the read-only store still rejects writes.
3. **Independent Exchange evidence.** Return Outlook to online mode. Observe
   that the account connection changes. Refresh the test account's web mailbox
   until that exact synthetic draft and its intended change appear, or record
   “server observation pending” at the end of a stated observation window.
   Record the exact observed item/change and time locally. The original tool
   receipt remains unverified. A Send/Receive start or a connected indicator
   does not verify the draft or mail delivery.
4. **Restart and separate profile.** Close Outlook normally while the MCP host
   stays open. Call a read against the selected store. Complete any prompts.
   For the same identifiable profile expect recovery or an actionable read
   timeout. Repeat with a different disposable profile chosen manually.
   An interrupted read must not be replayed into a different profile. Check
   readiness and explicitly select the intended authorized store for new work.
5. **Interrupted writes.** In the disposable store, exercise one synthetic send
   or invitation to the controlled recipient. If a prompt causes a timeout
   or the connection drops, retain the error, allow Outlook to settle and
   inspect Drafts, Calendar, Outbox, Sent Items and the recipient mailbox
   before considering another action. Count matching items. Where timing
   cannot be reproduced, record the scenario as unobserved. The deterministic
   queued and in-flight timeout cases run in the synthetic suite.

Record the commit, Windows/Outlook versions, scenario, expected/observed result
and any outstanding server observation. Publish only generic pass/fail summaries.

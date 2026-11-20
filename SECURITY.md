# Security development and verification

## Supported fork revisions

Security fixes target the current fork on `main`. Report issues against a
full fork commit, including for local installations reporting version 0.3.3.
Historical commits and upstream PyPI releases have no separate maintenance
branch here. See [installation and updates](docs/installation.md) and
[package provenance and release checks](docs/release.md).

## Repeat the automated checks

Use a fresh checkout and Python 3.10 or later. The workflow pins uv 0.13.9
and Ruff 0.17.6. From the repository root run:

```sh
uv sync --locked --extra dev
uv run --no-sync python -m pytest -q
uvx ruff@0.17.6 check .
node --test tests/ui-security.test.cjs
uv audit --locked
```

Pull requests targeting `main` and pushes to `main` run the complete
pytest suite. Windows runners exercise pywin32 STA initialization, reconnect
and write-replay handling with a simulated Outlook object. The STDIO suite
starts the real MCP SDK client and server with a synthetic bridge and a
temporary configuration. Linux also runs the platform-independent cases.
The JavaScript tests run the shipped view scripts with an in-memory host.
Fixtures should use temporary directories, invented subjects, and
`example.test` identities. Never copy an Outlook profile or a diagnostic
export into a fixture, Actions log or artifact.

When changing registrations, review the explicit STDIO inventory, metadata
and disabled-call assertions. For a new mailbox operation, exercise an
allowed store and a forbidden store, including the point at which a write
would occur. For calendar output, inspect both text representations and
structured data, private slots, missing/excluded items and error logs.
Treat email bodies, subjects, calendar text and filenames as external data.
Keep account and tool authorization in server code for every direct call.
A host with broad write permissions can still act on a misleading message.

The app accepts messages from its embedding parent and binds subsequent
traffic to the origin used for initialization. Opaque parent origins require
a wildcard delivery target. The embedding host and its sandbox policy are
part of the trust boundary. UI tests inspect rendered markup, outbound tool
calls and redacted state transitions. Host-specific browser behavior still
requires a separate integration check.

Ruff checks whitespace, invalid or undefined code and selected unsafe
execution patterns. The dependency job queries OSV for all registry
packages in the universal lockfile, including Windows and development
dependencies. On a finding, inspect the advisory, update the affected
package and lockfile, and rerun the tests and audit. A passing audit describes
the advisories known at query time. Package downloads and the advisory
query use the network during CI setup and review.

## Separate manual Windows and network verification

Arrange the scenario and allowed endpoints with the operator before running
this stage. Use a disposable Windows VM with classic Outlook, an isolated
synthetic profile, and a local STDIO client. Install dependencies before
capture. Create two synthetic stores, one read-only in the server policy,
and a private test appointment with a recognizable synthetic marker.

1. Start Windows Process Monitor with network events enabled. Filter by the
   process IDs of the server's Python process, the MCP host and OUTLOOK.EXE.
   Capture an idle Outlook baseline and record the time window.
2. Start the reviewed server locally with the synthetic policy. List tools
   and stores, read a normal appointment, then request the private item by
   list and ID under busy-only and exclusion policies. Restart the server
   after changing policy. Attempt a write to the read-only store and confirm
   in Outlook that the item is intact.
3. Compare each process's destinations and connection times against the
   baseline and the agreed endpoints. Attribute Outlook sync and host traffic
   separately from server traffic. Stop and investigate unexpected destinations.
   For a local-only scenario also repeat with outbound traffic blocked for the
   Python process and record any changes.
4. Inspect the displayed results and local diagnostic output for the synthetic
   marker. Keep raw captures, addresses, item IDs and logs local. Report only
   the reviewed commit, scenario, environment versions, pass/fail and any
   unexplained traffic. Dispose of the test profile and captures as agreed.

Keep this result separate from pytest and the simulated Windows bridge.
Static review cannot establish live COM behavior, runtime network destinations
or protection from a concurrent local process changing filesystem links.
A capture observes only the exercised paths and cannot establish absence of
all network traffic in other configurations.

## Reporting

If GitHub offers private reporting, use the repository's Security tab.
Otherwise [open a contact request](https://github.com/scepbjoern/Outlook-Classic-MCP/issues/new)
titled "Private security contact requested" with a way for the maintainer
to reach you privately. Share technical details only after that channel is
established. Include the affected fork commit and a minimal synthetic
description. Exclude account data, tokens and captures. The maintainer
triages reports on a best-effort basis and follows up privately about
reproduction and remediation. There is no guaranteed response deadline.

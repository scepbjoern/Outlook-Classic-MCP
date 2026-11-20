# Package verification and release

## Origin of version 0.3.3

The [PyPI provenance for both 0.3.3 files](https://pypi.org/project/outlook-classic-mcp/0.3.3/#files)
identifies the upstream GitHub repository, its `publish.yml` workflow and
commit [fec60243d54f13f7c43e321d2bcfc49c192ece0f](https://github.com/anasahmed07/Outlook-Classic-MCP/tree/fec60243d54f13f7c43e321d2bcfc49c192ece0f).
The November 2026 review downloaded both files without executing them and
matched their SHA-256 digests to PyPI's release metadata and attestation
subjects. The wheel's `outlook_mcp/server.py` differed from the fork at
`1240b00e898deede99215d37ed7514685f9beee7` after normalizing line endings.

| Upstream file | SHA-256 |
| --- | --- |
| `outlook_classic_mcp-0.3.3-py3-none-any.whl` | `df43d3169510e29a7c667a120a07124dd66b03ced219358540e97259d53fcf66` |
| `outlook_classic_mcp-0.3.3.tar.gz` | `0e216702e137c41d7834548b6415010feeeb0ac4f30c9ad71881933d6de1d9e9` |

The public version number identifies neither this fork nor a particular
fork commit. Use the [pinned source installation](installation.md).

## Rehearse without uploading

Start in a fresh, reviewed checkout with the pinned uv on PATH, an empty
or absent `dist` directory and the installed interpreter from the Windows
guide. Run each command only after the previous command succeeds.

```powershell
uv lock --check
uv build --offline --python .venv/Scripts/python.exe
.\.venv\Scripts\python.exe scripts/verify_dist.py dist
```

uv builds the wheel from the source archive with its bundled, pinned backend.
The inspection checks the exact package files in both archives against the
checkout, including the UI HTML. The wheel contains the server, entry point,
license and distribution metadata. Source-install scripts, client setup,
tests, and development tools live in Git. The PR workflow also installs the
wheel offline with locked runtime dependencies in a second environment,
checks dependency consistency, and loads its entry point without starting
Outlook. Record the Git commit, CI URL and SHA-256 values for the artifacts
from that run. Byte hashes may differ across checkouts with different line
endings. They identify the particular archives inspected.

## Publication decision

Local installations need no release tag, registry account or upload.
The **Verify release packages** workflow rehearses on PRs and manual runs.
A push to main or a tag does not request publication.

A maintainer choosing a future public fork release must first establish a
distinct fork-owned distribution name, update the package metadata and its
consumers, and configure PyPI Trusted Publishing for this repository,
`publish.yml` and the `pypi` environment. Set the repository variable
`PYPI_PROJECT` to that exact name. Protect the environment and release tags
and review the exact tagged commit and green CI before dispatching the
workflow on `v<package version>` with **publish** explicitly selected.

The gate rejects the upstream project name and an existing version. Only
PyPI's expected version-specific JSON 404 permits upload. Redirected errors,
network failures, rate limits, malformed responses and other statuses stop
the job. The publishing job receives only the archives verified in that run
and uses OIDC. After any future publication, compare the downloaded hashes,
publisher repository, workflow, and commit shown by PyPI against the review
record. A dry run cannot establish ownership or prove an upload will succeed.

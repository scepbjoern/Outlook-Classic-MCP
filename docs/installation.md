# Install or update the verified Windows fork

Start PowerShell in a directory you own. You need Git and uv 0.13.9.
If that uv version is already on PATH, skip the tool setup. With a
trusted Python installation, you can install the pinned uv wheel once in
a dedicated tools directory and select it for this terminal:

```powershell
$tools = Join-Path $env:LOCALAPPDATA 'outlook-mcp-tools-0.13.9'
python -m venv $tools
if ($LASTEXITCODE) { throw 'Tool environment creation failed' }
& "$tools/Scripts/python.exe" -m pip --isolated install --index-url https://pypi.org/simple --only-binary=:all: uv==0.13.9
if ($LASTEXITCODE) { throw 'uv installation failed' }
$env:Path = "$tools/Scripts;$env:Path"
uv --version
```

Alternatively obtain the Windows binary for **0.13.9** from the
[official uv release](https://github.com/astral-sh/uv/releases/tag/0.13.9).
Verify its published checksum before placing it on PATH. Git, uv, Python,
the package index and your machine form part of the installation trust boundary.

## Install one approved commit

Choose the full 40-character commit ID of an approved change **containing
this installer** from this fork. Compare the ID against its review and CI
results. A branch, inherited tag or package version alone does not identify
the reviewed code. Run the following in the same PowerShell window:

```powershell
$commit = (Read-Host 'Reviewed full fork commit').Trim()
if ($commit -notmatch '^[0-9a-fA-F]{40}$') { throw 'Full commit required' }
$install = Join-Path $PWD "outlook-$commit"
if (Test-Path -LiteralPath $install) { throw 'Choose a fresh directory' }
git clone --no-checkout https://github.com/scepbjoern/Outlook-Classic-MCP.git $install
if ($LASTEXITCODE) { throw 'Clone failed' }
git -C $install checkout --detach $commit
if ($LASTEXITCODE -or (git -C $install rev-parse HEAD) -ne $commit) { throw 'Commit mismatch' }
& "$install/install.bat"
if ($LASTEXITCODE) { throw 'Install failed' }
$serverPython = "$install/.venv/Scripts/python.exe"
& $serverPython -I -c "import sys, outlook_mcp; from importlib.metadata import version; print(sys.version, version('outlook-classic-mcp'), outlook_mcp.__file__)"
```

The bootstrap uses CPython 3.12.16 and the checked-in universal `uv.lock`.
It installs the local project as a non-editable package with its locked
runtime dependencies and checks the installed version. The pinned uv binary
includes the matching build backend. The package metadata and
`requirements.txt` list dependency constraints for other tools. Using these
ranges alone can select different versions. Keep the commit ID alongside
the environment path even when two commits report the same package version.

The check above only reads local files. To connect a client, use the
absolute `$serverPython` path as the command with arguments `-m outlook_mcp`.
Optional registration from PowerShell is
`& $serverPython "$install/scripts/install_to_clients.py"`.
For the optional agent plugin, select the local reviewed checkout and set
`OUTLOOK_MCP_PYTHON` to the same absolute executable path in the host's
environment before starting it. Hosts without variable expansion can use
the literal executable path in their STDIO configuration.

## Update and roll back

1. Review the new commit, its CI, and changes to the lockfile, tool policy
   and installation files. Run the same install block with the new SHA in
   a new directory. Leave the old environment in place.
2. Compare the new interpreter path, package version, and commit ID.
   Verify the intended local configuration and tool permissions. When
   ready, stop the client/server, change the executable path (or plugin
   variable), restart, and refresh tool discovery.
3. To roll back, stop the client/server and restore the old executable path
   and corresponding policy. Restart and refresh tool discovery again.
   Outlook operations already performed are independent of this change.

## Network boundaries

Git contacts GitHub for the fork. If the exact interpreter is missing, uv
can fetch the matching managed Python from GitHub release assets. Pip and
uv may contact PyPI and its file host for the tool and locked wheels.
A company proxy or configured mirror can change the observed destinations.
Installing or updating needs no mailbox, Outlook launch or client registration.

During operation the configured executable runs directly. The client and
server exchange JSON-RPC over local STDIO. Python calls local Outlook COM.
Outlook can sync with mail services and perform deliveries. A client can
forward returned tool data to its model provider or telemetry services.
See the [separate process-level network verification](../SECURITY.md#separate-manual-windows-and-network-verification)
for a capture procedure, its synthetic profile requirements and limits.

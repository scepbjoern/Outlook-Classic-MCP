"""Keep pytest isolated from a developer's Outlook configuration and application."""

import sys

import pytest


@pytest.fixture(autouse=True)
def synthetic_outlook_only(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    monkeypatch.setenv("XDG_CONFIG_HOME", str(tmp_path))
    monkeypatch.delenv("OUTLOOK_MCP_CONFIG", raising=False)
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_ANY_PATH", raising=False)

    def forbidden(*args, **kwargs):
        pytest.fail("Synthetic tests must supply their own Outlook object")

    monkeypatch.setattr("outlook_mcp.bridge._launch_outlook", forbidden)
    if sys.platform == "win32":
        monkeypatch.setattr("win32com.client.dynamic.Dispatch", forbidden)

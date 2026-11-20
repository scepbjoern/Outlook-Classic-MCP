"""Synthetic filesystem and Outlook attachment boundary checks."""

import os
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from outlook_mcp.client import mail
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir


@pytest.fixture
def profile(tmp_path, monkeypatch):
    root = tmp_path / "profile"
    root.mkdir()
    monkeypatch.setenv("USERPROFILE", str(root))
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_ANY_PATH", raising=False)
    return root


@pytest.mark.parametrize("validate", [validate_attachment_path, validate_output_dir])
@pytest.mark.parametrize("raw", [r"\\synthetic.invalid\share\file", "//synthetic.invalid/share/file",
                                 r"\\?\C:\fixture", r"\\.\fixture"])
def test_network_and_device_paths_rejected_before_io(monkeypatch, validate, raw):
    monkeypatch.delenv("OUTLOOK_MCP_ALLOW_ANY_PATH", raising=False)
    def unexpected(*args, **kwargs):
        pytest.fail("A rejected network or device name reached the filesystem")
    for method in ("resolve", "is_file", "mkdir"):
        monkeypatch.setattr(Path, method, unexpected)
    with pytest.raises(OutlookError):
        validate(raw)


@pytest.mark.parametrize("validate", [validate_attachment_path, validate_output_dir])
def test_traversal_rejected_before_creating_or_reading(profile, validate):
    with pytest.raises(OutlookError):
        validate(str(profile / ".." / "outside"))
    assert not (profile.parent / "outside").exists()


@pytest.mark.parametrize("kind", ["symlink", "junction"])
def test_link_cannot_escape_profile(profile, kind):
    outside = profile.parent / "outside"
    outside.mkdir()
    (outside / "synthetic.txt").write_text("synthetic")
    link = profile / "linked"
    if kind == "junction":
        if os.name != "nt":
            pytest.skip("Windows directory junction")
        from _winapi import CreateJunction
        CreateJunction(str(outside), str(link))
    else:
        try:
            link.symlink_to(outside, target_is_directory=True)
        except OSError as exc:
            if getattr(exc, "winerror", None) == 1314:
                pytest.skip("Windows symlink privilege unavailable; junction tested separately")
            raise
    with pytest.raises(OutlookError, match="outside the user profile"):
        validate_attachment_path(str(link / "synthetic.txt"))
    with pytest.raises(OutlookError, match="outside the user profile"):
        validate_output_dir(str(link / "new"))
    assert not (outside / "new").exists()


def save(profile, attachment, monkeypatch):
    item = SimpleNamespace(Attachments=[attachment])
    monkeypatch.setattr(mail, "get_item_by_id", lambda *args: item)
    return mail.save_attachments(None, None, entry_id="synthetic", output_dir=str(profile))


@pytest.mark.parametrize("name", ["..", "../file", r"folder\file", "report:stream",
                                  "CON.txt", "nul .txt", "LPT¹.pdf", "data. ", "CONOUT$"])
def test_unsafe_sender_filename_never_reaches_save(profile, monkeypatch, name):
    att = SimpleNamespace(FileName=name, SaveAsFile=Mock())
    with pytest.raises(OutlookError):
        save(profile, att, monkeypatch)
    att.SaveAsFile.assert_not_called()


def test_colliding_names_preserve_existing_file(profile, monkeypatch):
    existing = profile / "synthetic.txt"
    existing.write_text("original")
    att = SimpleNamespace(FileName=existing.name, SaveAsFile=Mock())
    result = save(profile, att, monkeypatch)
    target = str(profile / "synthetic (1).txt")
    att.SaveAsFile.assert_called_once_with(target)
    assert result["files"] == [target]
    assert existing.read_text() == "original"

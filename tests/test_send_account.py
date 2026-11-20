"""Sender binding and signature reloads with synthetic delivery stores."""
from unittest.mock import MagicMock

import pytest

from outlook_mcp.errors import WriteOutcomeUnknown

from outlook_mcp.client.mail import send_mail
from outlook_mcp.client.mailbox import Access, MailboxProfile
from outlook_mcp.config import AccountPolicy
from outlook_mcp.constants import OL_CLASS_MAIL
from tests.test_mailbox import make_profile


def test_send_mail_with_account_saves_into_account_drafts():
    ns = make_profile()
    policy = AccountPolicy(write=frozenset({"b@example.test"}))
    target = MailboxProfile(ns, policy).resolve("B@example.test", None, Access.WRITE)
    result = send_mail(ns, target, to=["guest@example.test"], subject="Synthetic",
                       body="Body", save_only=True)
    draft = target.store.GetDefaultFolder.return_value.Items.Add.return_value
    assert draft.SendUsingAccount is ns.Accounts[1]
    assert result["store_id"] == "store-b"
    draft.Save.assert_called_once()
    ns.Stores[0].GetDefaultFolder.assert_not_called()


def test_single_account_signature_closes_inspector_and_sends_reloaded_item():
    ns = make_profile(("a",))
    target = MailboxProfile(ns, None).resolve(None, None, Access.WRITE)
    draft = target.store.GetDefaultFolder.return_value.Items.Add.return_value
    draft.HTMLBody = "<html><body><p>SIG</p></body></html>"
    reloaded = MagicMock()
    reloaded.Parent = draft.Parent
    reloaded.Class = OL_CLASS_MAIL
    ns.GetItemFromID.side_effect = None
    ns.GetItemFromID.return_value = reloaded

    send_mail(ns, target, to=["guest@example.test"], subject="Synthetic",
              body="Hello", signature=True)
    draft.Save.assert_called_once()
    draft.GetInspector.Close.assert_called_once_with(0)
    draft.Send.assert_not_called()
    ns.GetItemFromID.assert_called_once_with(draft.EntryID, target.store_id)
    assert reloaded.SendUsingAccount is ns.Accounts[0]
    reloaded.Send.assert_called_once()
    assert "Hello" in draft.HTMLBody and "SIG" in draft.HTMLBody


def test_signature_reload_failure_reports_existing_local_draft():
    ns = make_profile(("a",))
    target = MailboxProfile(ns, None).resolve(None, None, Access.WRITE)
    draft = target.store.GetDefaultFolder.return_value.Items.Add.return_value
    draft.HTMLBody = ""
    ns.GetItemFromID.side_effect = None
    ns.GetItemFromID.return_value = None
    with pytest.raises(WriteOutcomeUnknown, match="saved a draft locally.*Do not repeat"):
        send_mail(ns, target, to=["guest@example.test"], subject="Synthetic",
                  body="Example", signature=True)
    draft.Save.assert_called_once()
    draft.Send.assert_not_called()

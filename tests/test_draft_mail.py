"""Check COM effects behind the registered draft tool using two synthetic stores."""
from types import SimpleNamespace
from unittest.mock import PropertyMock

import pytest

from outlook_mcp.constants import OL_FOLDER_DRAFTS, OL_MAIL_ITEM
from tests.test_mailbox import make_profile, registered


MESSAGE = {"to": ["guest@example.test"], "subject": "Synthetic draft", "body": "Review <this>"}


@pytest.fixture
def draft_stores():
    ns = make_profile()
    yield ns
    for store in ns.Stores:
        store.GetDefaultFolder.return_value.Items.Add.return_value.Send.assert_not_called()
    ns.GetDefaultFolder.assert_not_called()
    ns.GetItemFromID.assert_not_called()


@pytest.mark.parametrize("signature", [False, True])
def test_draft_composition_and_local_save(draft_stores, signature):
    ns = draft_stores
    primary, selected = ns.Stores
    folder = selected.GetDefaultFolder.return_value
    mail = folder.Items.Add.return_value
    mail.HTMLBody = "<body>Signature</body>"

    result = registered(ns, read=())(
        "draft_mail", account="b@example.test", **MESSAGE, signature=signature,
        cc=["cc@example.test"], bcc=["bcc@example.test"],
    )

    primary.GetDefaultFolder.assert_not_called()
    selected.GetDefaultFolder.assert_called_once_with(OL_FOLDER_DRAFTS)
    folder.Items.Add.assert_called_once_with(OL_MAIL_ITEM)
    assert mail.SendUsingAccount is ns.Accounts[1]
    assert (mail.To, mail.CC, mail.BCC) == ("guest@example.test", "cc@example.test", "bcc@example.test")
    if signature:
        assert mail.HTMLBody == "<body>Review &lt;this&gt;<br><br>Signature</body>"
        mail.GetInspector.Close.assert_called_once_with(0)
    else:
        assert mail.Body == MESSAGE["body"]
    mail.Save.assert_called_once_with()
    assert result == {
        "status": "saved_to_drafts", "local_outcome": "saved", "exchange_sync": "unverified",
        "entry_id": mail.EntryID, "store_id": "store-b", "subject": MESSAGE["subject"],
    }


@pytest.mark.parametrize(("returned_address", "returned_store", "accepted"), [
    ("b@example.test", 1, True),
    ("a@example.test", 1, False),
    ("b@example.test", 0, False),
    (None, None, False),
])
def test_late_bound_sender_needs_verified_object_reference(
    draft_stores, monkeypatch, returned_address, returned_store, accepted,
):
    ns = draft_stores
    folder = ns.Stores[1].GetDefaultFolder.return_value
    mail = folder.Items.Add.return_value
    sender = ns.Accounts[0]
    def put_reference(property_id, locale, flags, wants_result, account):
        nonlocal sender
        assert (property_id, locale, flags, wants_result) == (64209, 0, 8, 0)
        assert account is ns.Accounts[1]
        sender = None if returned_store is None else SimpleNamespace(
            SmtpAddress=returned_address, DeliveryStore=ns.Stores[returned_store],
        )

    # The ordinary property assignment succeeds but the getter still returns
    # the primary account until Outlook accepts an explicit reference put.
    prop = PropertyMock(side_effect=lambda *args: None if args else sender)
    monkeypatch.setattr(type(mail), "SendUsingAccount", prop, raising=False)
    dispatch = mail._oleobj_
    dispatch.GetIDsOfNames.return_value = 64209
    dispatch.Invoke.side_effect = put_reference
    invoke = registered(ns)
    if accepted:
        assert invoke("draft_mail", store_id="store-b", **MESSAGE)["local_outcome"] == "saved"
        mail.Save.assert_called_once_with()
    else:
        with pytest.raises(Exception, match="did not bind"):
            invoke("draft_mail", store_id="store-b", **MESSAGE)
        mail.Save.assert_not_called()
    folder.Items.Add.assert_called_once_with(OL_MAIL_ITEM)
    dispatch.GetIDsOfNames.assert_called_once_with("SendUsingAccount")
    dispatch.Invoke.assert_called_once()


def test_read_only_store_rejected_before_draft_creation(draft_stores):
    with pytest.raises(Exception, match="does not grant write"):
        registered(draft_stores)("draft_mail", store_id="store-a", **MESSAGE)
    for store in draft_stores.Stores:
        store.GetDefaultFolder.assert_not_called()


@pytest.mark.parametrize(("failure", "error"), [
    ("wrong_parent", "outside the selected store"),
    ("attachment", "Attachment not found"),
    ("save", "local draft save could not be confirmed.*Do not repeat"),
    ("inspector_close", "saved a draft locally.*Do not repeat"),
])
def test_draft_failure_does_not_dispatch_or_repeat(draft_stores, failure, error, tmp_path, monkeypatch):
    ns = draft_stores
    primary, selected = ns.Stores
    mail = selected.GetDefaultFolder.return_value.Items.Add.return_value
    mail.HTMLBody = "<body>Signature</body>"
    options = {}
    if failure == "wrong_parent":
        mail.Parent = primary.GetDefaultFolder.return_value
    elif failure == "attachment":
        monkeypatch.setenv("USERPROFILE", str(tmp_path))
        options["attachments"] = [str(tmp_path / "missing.txt")]
    elif failure == "save":
        mail.Save.side_effect = RuntimeError("synthetic save failure")
    else:
        options["signature"] = True
        mail.GetInspector.Close.side_effect = RuntimeError("synthetic inspector failure")

    with pytest.raises(Exception, match=error):
        registered(ns)("draft_mail", store_id="store-b", **MESSAGE, **options)

    assert mail.Save.call_count == int(failure in {"save", "inspector_close"})
    selected.GetDefaultFolder.assert_called_once_with(OL_FOLDER_DRAFTS)
    primary.GetDefaultFolder.assert_not_called()

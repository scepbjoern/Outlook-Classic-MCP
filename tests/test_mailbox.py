"""Exercise the registered server against two synthetic Exchange stores."""
import asyncio
import json
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from outlook_mcp.client.mailbox import Access, MailboxProfile
from outlook_mcp.config import AccountPolicy, ToolConfigError, parse_tool_policy
from outlook_mcp.constants import OL_CLASS_APPOINTMENT, OL_FOLDER_CALENDAR
from outlook_mcp.errors import OutlookError
from outlook_mcp.server import build_server


class Collection(list):
    @property
    def Count(self):
        return len(self)

    def Item(self, index):
        return self[index - 1]


def make_profile(names=("a", "b")):
    namespace = MagicMock()
    accounts = Collection()
    for name in names:
        store = MagicMock()
        store.StoreID, store.DisplayName = "store-" + name, name
        store.ExchangeStoreType = 4
        folder = store.GetDefaultFolder.return_value
        folder.StoreID = store.StoreID
        event = folder.Items.Add.return_value
        event.Parent = folder
        event.Class = OL_CLASS_APPOINTMENT
        event.EntryID = "event-" + name
        event.Subject = name
        event.Body = ""
        event.Categories = ""
        event.ReminderMinutesBeforeStart = 15
        event.Location = ""
        event.Organizer = ""
        event.Recipients.Count = 0
        event.Recipients.__bool__.return_value = False
        event.RecurrenceState = 0
        event.AllDayEvent = False
        event.Start = event.End = None
        folder.Items.Restrict.return_value = [event]
        accounts.append(SimpleNamespace(SmtpAddress=name+"@example.test", DeliveryStore=store,
                                        UserName=name, DisplayName=name, AccountType=0))
    namespace.Accounts = accounts
    namespace.Stores = Collection(account.DeliveryStore for account in accounts)
    def lookup(entry_id, store_id):
        for store in namespace.Stores:
            event = store.GetDefaultFolder.return_value.Items.Add.return_value
            if store.StoreID == store_id and event.EntryID == entry_id:
                return event
        return None
    namespace.GetItemFromID.side_effect = lookup
    return namespace


def registered(namespace, *, read=("a", "b"), write=("b",)):
    def addresses(names):
        return [name+"@example.test" for name in names]
    policy = parse_tool_policy(json.dumps({"version": 2, "accounts": {
        "read": addresses(read), "write": addresses(write),
    }}))
    server, bridge = build_server(policy=policy)
    async def synthetic_call(function, *args, retry_on_disconnect=True, **kwargs):
        return function(None, namespace, *args, **kwargs)
    bridge.call = synthetic_call
    def invoke(tool_name, **args):
        result = asyncio.run(server.call_tool("outlook_" + tool_name, args))
        if hasattr(result, "structuredContent"):
            return result.structuredContent
        content = result[0] if isinstance(result, tuple) else result
        return json.loads(content[0].text)
    return invoke


def test_registered_calendar_routes_each_operation_to_selected_store():
    ns = make_profile()
    a, b = ns.Stores
    call = registered(ns)
    assert call("list_events", account="A@example.test")["items"][0]["entry_id"] == "event-a"
    assert call("list_events", store_id=b.StoreID)["items"][0]["entry_id"] == "event-b"
    assert call("get_event", entry_id="event-b", store_id=b.StoreID)["store_id"] == b.StoreID
    a.reset_mock()
    result = call("create_event", account="b@example.test", subject="Synthetic meeting",
                  start="2026-11-21T10:00:00", end="2026-11-21T11:00:00",
                  attendees=["guest@example.test"])
    assert result["store_id"] == b.StoreID
    b.GetDefaultFolder.assert_called_with(OL_FOLDER_CALENDAR)
    event = b.GetDefaultFolder.return_value.Items.Add.return_value
    event.Save.assert_called_once()
    assert event.SendUsingAccount is ns.Accounts[1]
    event.Send.assert_called_once()
    call("update_event", entry_id=event.EntryID, subject="Changed", store_id=b.StoreID)
    assert event.Subject == "Changed"
    call("respond_event", entry_id=event.EntryID, response="accept", store_id=b.StoreID)
    response = event.Respond.return_value
    assert response.SendUsingAccount is ns.Accounts[1]
    response.Send.assert_called_once()
    call("delete_event", entry_id=event.EntryID, store_id=b.StoreID)
    event.Delete.assert_called_once()
    assert not a.mock_calls
    ns.GetDefaultFolder.assert_not_called()


@pytest.mark.parametrize(("selectors", "error"), [
    ({}, "Specify account or store_id"),
    ({"account":"missing@example.test"}, "Unknown"),
    ({"store_id":"missing"}, "Unknown"),
    ({"account":"b@example.test", "store_id":"store-a"}, "different"),
    ({"store_id":"store-a"}, "does not grant"),
])
def test_registered_write_rejected_before_item_lookup(selectors, error):
    ns = make_profile()
    call = registered(ns)
    with pytest.raises(Exception, match=error):
        call("delete_event", entry_id="event-a", **selectors)
    ns.GetItemFromID.assert_not_called()
    for store in ns.Stores:
        store.GetDefaultFolder.assert_not_called()


@pytest.mark.parametrize("tool,extra", [
    ("get_event", {"response_format":"json"}), ("update_event", {"subject":"blocked"}),
    ("delete_event", {}), ("respond_event", {"response":"decline"}),
    ("set_category", {"categories":"blocked"}),
])
def test_provider_returning_foreign_item_is_rejected(tool, extra):
    ns = make_profile()
    event = ns.Stores[0].GetDefaultFolder.return_value.Items.Add.return_value
    ns.GetItemFromID.side_effect = None
    ns.GetItemFromID.return_value = event
    with pytest.raises(Exception, match="outside the selected store"):
        registered(ns)(tool, entry_id=event.EntryID, store_id="store-b", **extra)
    ns.GetItemFromID.assert_called_once_with(event.EntryID, "store-b")
    assert event.Subject == "a"
    event.Delete.assert_not_called()
    event.Respond.assert_not_called()
    event.Save.assert_not_called()


def test_independent_permissions_and_filtered_identity():
    ns = make_profile()
    call = registered(ns, read=(), write=("b",))
    identity = call("whoami", response_format="json")
    assert [row["smtp_address"] for row in identity["accounts"]] == ["b@example.test"]
    assert identity["current_user"] is None and identity["accounts"][0]["write"]
    with pytest.raises(Exception, match="does not grant read"):
        call("list_events", store_id="store-b")
    ns.Stores[1].GetDefaultFolder.assert_not_called()
    assert call("create_event", store_id="store-b", subject="Write only",
                start="2026-11-21", end="2026-11-22")["status"] == "created"


@pytest.mark.parametrize("defect", ["duplicate_address", "shared_store", "no_store"])
def test_ambiguous_or_missing_identity(defect):
    ns = make_profile()
    if defect == "duplicate_address":
        ns.Accounts[0].SmtpAddress = ns.Accounts[1].SmtpAddress
    elif defect == "shared_store":
        ns.Accounts[0].DeliveryStore = ns.Accounts[1].DeliveryStore
    else:
        ns.Accounts[1].DeliveryStore = None
    with pytest.raises(OutlookError, match="Ambiguous|no usable"):
        MailboxProfile(ns, AccountPolicy()).resolve("b@example.test", None, Access.WRITE)


def test_legacy_rights_follow_actual_account_count():
    ns = make_profile(("a",))
    mailbox = MailboxProfile(ns, None).resolve(None, None, Access.READ_WRITE)
    assert mailbox.store is ns.Stores[0]
    with pytest.raises(OutlookError, match="does not grant"):
        MailboxProfile(ns, AccountPolicy()).resolve(None, None, Access.READ)
    with pytest.raises(OutlookError, match="version 2"):
        MailboxProfile(make_profile(), None).resolve(None, "store-a", Access.READ)


def test_wrong_store_lookup_never_retries_entry_id_alone():
    ns = make_profile()
    with pytest.raises(Exception, match="Item not found in the selected store"):
        registered(ns)("update_event", store_id="store-b", entry_id="event-a", subject="Wrong store")
    ns.GetItemFromID.assert_called_once_with("event-a", "store-b")
    for store in ns.Stores:
        store.GetDefaultFolder.return_value.Items.Add.return_value.Save.assert_not_called()


def test_identical_entry_ids_are_independent_between_stores():
    ns = make_profile()
    for store in ns.Stores:
        store.GetDefaultFolder.return_value.Items.Add.return_value.EntryID = "same"
    call = registered(ns)
    for name in ("a", "b"):
        result = call("get_event", entry_id="same", store_id="store-" + name)
        assert result["subject"] == name


@pytest.mark.parametrize("raw", [
    '{"accounts":{}}', '{"version":1,"accounts":{}}', '{"version":2,"accounts":null}',
    '{"version":2,"accounts":{"reads":[]}}', '{"version":2,"accounts":{"read":"a@example.test"}}',
    '{"version":2,"accounts":{"read":["Alice"]}}', '{"version":2,"accounts":{"write":[null]}}',
    '{"version":2,"accounts":{"read":["a@example.test","A@example.test"]}}',
    '{"version":2,"accounts":{"read":[],"read":[]}}',
])
def test_invalid_account_policy(raw):
    with pytest.raises(ToolConfigError):
        parse_tool_policy(raw)


def test_non_calendar_operations_use_selected_store():
    ns = make_profile()
    a, b = ns.Stores
    call = registered(ns)
    selector = {"store_id": b.StoreID}
    folder = b.GetDefaultFolder.return_value
    folder.Name = "Inbox"
    root = b.GetRootFolder.return_value
    root.StoreID, root.Name = b.StoreID, "Mailbox B"
    root.Folders = [folder]
    folder.Folders = []
    folder.Items.Count = root.Items.Count = 0
    folder.UnReadItemCount = root.UnReadItemCount = 0
    root.DefaultItemType = folder.DefaultItemType = 0
    folders = call("list_folders", response_format="json", **selector)
    assert {row["path"] for row in folders["items"]} == {"Mailbox B", "Mailbox B/Inbox"}
    assert call("list_mails", response_format="json", **selector)["folder"] == "Inbox"
    assert call("list_contacts", response_format="json", **selector)["store_id"] == b.StoreID
    assert call("create_task", subject="Synthetic task", **selector)["store_id"] == b.StoreID
    b.Categories.Count = 0
    assert call("list_categories", response_format="json", **selector)["count"] == 0
    b.GetRules.return_value.Count = 0
    assert call("list_rules", response_format="json", **selector)["count"] == 0
    b.PropertyAccessor.GetProperty.return_value = True
    assert call("get_out_of_office", response_format="json", **selector)["out_of_office"]
    assert not a.mock_calls
    ns.GetDefaultFolder.assert_not_called()


def test_unsupported_provider_operations_stop_before_profile_lookup():
    ns = make_profile()
    store = ns.Stores[1]
    store.ExchangeStoreType = 3
    call = registered(ns)
    with pytest.raises(Exception, match="single-account"):
        call("list_categories", store_id=store.StoreID)
    with pytest.raises(Exception, match="single-account"):
        call("resolve_name", name="Synthetic", store_id=store.StoreID)
    ns.CreateRecipient.assert_not_called()
    store.Categories.Item.assert_not_called()


def test_rule_with_foreign_folder_is_rejected_before_save():
    ns = make_profile()
    rules = ns.Stores[1].GetRules.return_value
    rule = rules.Item.return_value
    rules.Count = 1
    rules.__iter__.return_value = iter([rule])
    rule.Name, rule.Enabled = "Synthetic rule", False
    rule.Actions = [SimpleNamespace(ActionType=1, Enabled=True, Folder=ns.Stores[0])]
    with pytest.raises(Exception, match="outside the selected store"):
        registered(ns)("toggle_rule", store_id="store-b", rule_name=rule.Name, enabled=True)
    assert not rule.Enabled
    rules.Save.assert_not_called()


def test_exchange_directory_reads_only_entries_of_matching_store():
    from outlook_mcp.client.contacts import _directory_for_store

    ns = make_profile()
    mailbox = MailboxProfile(ns, AccountPolicy(read=frozenset({"b@example.test"}))).resolve(
        None, "store-b", Access.READ,
    )
    lists = []
    for store in ns.Stores:
        provider = store.PropertyAccessor
        provider.BinaryToString.side_effect = lambda uid: uid
        provider.GetProperty.return_value = store.StoreID
        lists.append(SimpleNamespace(AddressListType=0, PropertyAccessor=provider))
    ns.AddressLists = lists
    assert _directory_for_store(ns, mailbox) is lists[1]
    ns.AddressLists = [lists[0]]
    assert _directory_for_store(ns, mailbox) is None
    ns.AddressLists = [lists[1], lists[1]]
    assert _directory_for_store(ns, mailbox) is None
    ns.GetGlobalAddressList.assert_not_called()

"""Store-specific category assignments through the registered calendar tools."""

import datetime as dt
from types import SimpleNamespace

import pytest

from outlook_mcp.client import categories
from tests.test_calendar_recurrence import occurrence, selected, series
from tests.test_calendar_sensitivity import EXISTING, NEW_EVENT, STORE
from tests.test_mailbox import Collection, WINDOW, make_profile, registered


LABEL = "Révision, été – 東京"
OTHER = "Prüfung (A/B)!"
LEGACY = "Archive 旧"


def catalog(*names):
    return Collection(SimpleNamespace(Name=name, Color=index) for index, name in enumerate(names, 1))


def mock_keywords(event, separator):
    def read(schema):
        assert schema == "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Keywords"
        return tuple(name.strip() for name in event.Categories.split(separator))
    event.PropertyAccessor.GetProperty.side_effect = read


@pytest.fixture
def profile(monkeypatch):
    monkeypatch.setattr(categories, "_list_separator", lambda: ";")
    ns = make_profile()
    a, b = ns.Stores
    a.Categories = catalog("Atelier")
    b.Categories = catalog(LABEL, OTHER, "Privat")
    mock_keywords(b.GetDefaultFolder.return_value.Items.Add.return_value, ";")
    return ns


def test_create_and_list_use_the_selected_master_list(profile):
    call = registered(profile, write=("a", "b"))
    for store, names in zip(profile.Stores, [["Atelier"], [LABEL, OTHER, "Privat"]]):
        selector = {"account": store.DisplayName + "@example.test"}
        assert call("list_categories", **selector, response_format="json") == {
            "count": len(names), "items": [{"name": n, "color": i} for i, n in enumerate(names, 1)]}
        receipt = call("create_event", **selector, **NEW_EVENT, categories=names + names)
        detail = call("get_event", entry_id=receipt["entry_id"], store_id=receipt["store_id"])
        assert detail["categories"] == ";".join(names)
        assert detail["sensitivity"] == 0


def test_delta_preserves_legacy_names_and_replace_explicitly_clears(profile):
    ev = profile.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    ev.Categories = f"{LABEL}; {OTHER};{LEGACY}"
    call = registered(profile)
    call("update_event", **EXISTING, subject="Nouvel atelier")
    assert ev.Categories == f"{LABEL}; {OTHER};{LEGACY}"
    for _ in range(2):
        call("update_event", **EXISTING, add_categories=["Privat", "Privat"], remove_categories=[LABEL])
        assert ev.Categories == f"{OTHER};{LEGACY};Privat"
    call("update_event", **EXISTING, replace_categories=[LABEL])
    assert ev.Categories == LABEL
    call("update_event", **EXISTING, replace_categories=[])
    assert ev.Categories == ""
    assert ev.Sensitivity == 0


@pytest.mark.parametrize("tool,fields,error", [
    ("create_event", {"categories": [LABEL, "Atelier"], "attendees": ["guest@example.test"]}, "Unknown category"),
    ("update_event", {"replace_categories": [OTHER, "Atelier"]}, "Unknown category"),
    ("update_event", {"remove_categories": ["Atelier"], "add_categories": [LABEL]}, "Unknown category"),
    ("update_event", {"add_categories": [LABEL], "remove_categories": [LABEL]}, "both added and removed"),
    ("update_event", {"replace_categories": [], "add_categories": []}, "replace_categories alone"),
    ("update_event", {"add_categories": ["Séance;Bleu"]}, "separator"),
])
def test_invalid_requests_rejected_before_any_mutation(profile, tool, fields, error):
    store = profile.Stores[1]
    store.Categories = catalog(LABEL, OTHER, "Séance;Bleu")
    items = store.GetDefaultFolder.return_value.Items
    ev = items.Add.return_value
    ev.Categories = LEGACY
    args = {**NEW_EVENT, **STORE} if tool == "create_event" else {**EXISTING, "subject": "Changed", "sensitivity": 2}
    with pytest.raises(Exception, match=error):
        registered(profile)(tool, **args, **fields)
    assert (ev.Subject, ev.Categories, ev.Sensitivity) == ("b", LEGACY, 0)
    items.Add.assert_not_called()
    ev.Save.assert_not_called()
    ev.Send.assert_not_called()


@pytest.mark.parametrize("tool,args,fields", [
    ("create_event", {**NEW_EVENT, **STORE}, {"categories": [LABEL]}),
    ("update_event", EXISTING, {"add_categories": [LABEL]}),
])
def test_category_validation_uses_listing_provider_guard_only_when_needed(profile, tool, args, fields):
    profile.Stores[1].ExchangeStoreType = 3
    call = registered(profile)
    with pytest.raises(Exception, match="single-account, single-store"):
        call(tool, **args, **fields)
    call(tool, **args)


@pytest.mark.parametrize("privacy", ["busy_only", "exclude"])
def test_private_category_receipts_and_reads(profile, privacy):
    call = registered(profile, privacy={"default": privacy})
    created = call("create_event", **STORE, **NEW_EVENT, categories=["Privat"], sensitivity=2)
    assert set(created) == ({"status", "subject", "start", "end", "busy_status"} if privacy == "busy_only" else {"status"})
    assert call("update_event", **EXISTING, remove_categories=["Privat"], add_categories=[LABEL]) == {"status": "updated"}
    rows = call("list_events", **STORE, **WINDOW)["items"]
    if privacy == "exclude":
        assert rows == []
        with pytest.raises(Exception, match="Item not found"):
            call("get_event", **EXISTING)
    else:
        detail = call("get_event", **EXISTING)
        assert rows == [detail] and set(detail) == {"subject", "start", "end", "busy_status"}
        assert detail["subject"] == "Appointment"
    ev = profile.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    assert (ev.Categories, ev.Sensitivity) == (LABEL, 2)


def test_update_categories_on_middle_occurrence_only(monkeypatch):
    monkeypatch.setattr(categories, "_list_separator", lambda: ",")
    ns, master, pattern = series()
    label = "Prévu; 東京"
    ns.Stores[1].Categories = catalog(label, OTHER)
    master.Categories = OTHER
    events = [occurrence(master, dt.datetime(2026, 11, day, 8, tzinfo=dt.timezone.utc)) for day in (22, 23, 24)]
    for event in events:
        event.Categories = OTHER
        mock_keywords(event, ",")
    master.Parent.Items.Restrict.return_value = events
    pattern.GetOccurrence.return_value = events[1]
    call = registered(ns)
    rows = call("list_events", **STORE, **WINDOW)["items"]
    call("update_event", **selected(rows[1]), add_categories=[label])
    assert call("get_event", **selected(rows[1]))["categories"] == f"{OTHER},{label}"
    events[1].Save.assert_called_once_with()
    for untouched in (master, events[0], events[2]):
        assert untouched.Categories == OTHER
        untouched.Save.assert_not_called()


def test_legacy_set_category_accepts_unlisted_names_and_empty_string(profile):
    call = registered(profile)
    for raw in ("Unlisted, Fictional", ""):
        assert call("set_category", **EXISTING, categories=raw)["categories"] == raw


@pytest.mark.parametrize("keywords", [("Ancien; nom",), None, "Ancien; nom"])
def test_unrepresentable_or_unreadable_existing_labels_prevent_partial_edit(profile, keywords):
    ev = profile.Stores[1].GetDefaultFolder.return_value.Items.Add.return_value
    ev.Categories = "Ancien; nom"
    ev.PropertyAccessor.GetProperty.side_effect = None
    ev.PropertyAccessor.GetProperty.return_value = keywords
    with pytest.raises(Exception, match="separator|Cannot read"):
        registered(profile)("update_event", **EXISTING, subject="Changed", add_categories=[LABEL])
    assert (ev.Subject, ev.Categories) == ("b", "Ancien; nom")
    ev.Save.assert_not_called()

import json

import pytest

from src import extractor, parser, profiles
from src.profiles import ColumnRef, Profile, SheetLayout, SourceSchema, lookup_group
from tests.conftest import make_report_pdf


def _profile(**groups):
    return Profile(profile_name="P", group_to_sheet=dict(groups))


def test_group_lookup_normalises_safely():
    p = _profile(**{"NORTH MARKET": "S1", "Beta  Two.": "S2"})
    assert lookup_group(p, "north market").sheet == "S1"              # case
    assert lookup_group(p, "BETA TWO").sheet == "S2"               # punctuation / spaces
    alias = lookup_group(p, "NORTHMARKET")                            # space-insensitive
    assert (alias.status, alias.sheet, alias.saved_name) == ("alias", "S1", "NORTH MARKET")
    assert lookup_group(p, "NORTH MARKT").status == "unmapped"      # no fuzzy matching


def test_ambiguous_alias_is_not_guessed_and_is_settled_once():
    p = _profile(**{"EAST END": "S1", "EASTE ND": "S2"})
    match = lookup_group(p, "EASTEND")
    assert match.status == "ambiguous" and match.candidates == ["S1", "S2"]
    profiles.remember_group(p, "EASTEND", "S1")
    assert lookup_group(p, "EASTEND").sheet == "S1"


def test_ignored_group():
    p = _profile(A="S1")
    profiles.ignore_group(p, "Walk In")
    assert lookup_group(p, "WALK IN").status == "ignored"


def test_store_roundtrip_rename_delete_import_export(store):
    p = Profile(profile_name="Sales", source=SourceSchema(date_field="Doc Date"),
                group_to_sheet={"A": "S1"},
                sheet_template=SheetLayout(1, ColumnRef("A", "DATE"), ColumnRef("B", "CHALLAN")),
                sheet_overrides={"S9": SheetLayout(2, ColumnRef("A"), ColumnRef("C", "KARKHANA"))})
    store.save(p)
    loaded = store.load("Sales")
    assert loaded.sheet_overrides["S9"].amount.letter == "C"
    assert loaded.profile_version == profiles.PROFILE_VERSION and loaded.created
    store.rename("Sales", "Monthly sales")
    assert store.names() == ["Monthly sales"]
    exported = profiles.to_json(store.load("Monthly sales"))
    imported = store.import_json(exported)
    assert imported.profile_name == "Monthly sales (2)"
    assert store.delete("Monthly sales") and store.names() == ["Monthly sales (2)"]


def test_legacy_mapping_json_import_drops_positional_fields(store):
    legacy = {"source": {"group_label": "Customer Name", "date_field": "col_1", "amount_field": "col_2"},
              "group_to_sheet": {"A": "S1"},
              "sheets": {"S1": {"header_row": 1, "date_target_mode": "column_letter", "date_column": "A",
                                "amount_target_mode": "column_letter", "amount_column": "B"}},
              "write_rules": {"write_action": "replace", "output_type": "numeric"}}
    p = store.import_json(json.dumps(legacy))
    assert p.source.date_field == "" and p.source.amount_field == ""   # re-mapped by label later
    assert p.group_to_sheet == {"A": "S1"} and p.sheet_template.amount.letter == "B"


def _report(sample_groups):
    return parser.parse(extractor.extract(make_report_pdf(sample_groups), "r.pdf"))


def test_fields_restored_by_label_not_position(sample_groups):
    report = _report(sample_groups)
    good = profiles.suggest_profile(report)
    assert (good.source.date_field, good.source.amount_field) == ("Doc Date", "Total Amount")
    assert profiles.field_issues(good, report) == {}

    # A saved date field pointing at the reference column is refused, not swapped.
    wrong = profiles.suggest_profile(report)
    wrong.source.date_field = "Ref No"
    issues = profiles.field_issues(wrong, report)
    assert "date_field" in issues and "text" in issues["date_field"]

    # A saved field that no longer exists is flagged; nothing is substituted.
    missing = profiles.suggest_profile(report)
    missing.source.date_field = "Invoice Date"
    assert "not in this document" in profiles.field_issues(missing, report)["date_field"]


def test_recognised_by_structure_most_recent_success_wins(store, sample_groups):
    report = _report(sample_groups)
    older = profiles.suggest_profile(report, "Older")
    newer = profiles.suggest_profile(report, "Newer")
    other = Profile(profile_name="Other report", source=SourceSchema(
        group_label="Supplier", date_field="Date", amount_field="Amount"))
    for p in (older, newer, other):
        store.save(p)
    store.mark_success(store.load("Older"))
    store.mark_success(store.load("Newer"))
    assert store.find_for(report).profile_name == "Newer"
    assert store.last_successful().profile_name == "Newer"


def test_newer_profile_versions_are_rejected():
    with pytest.raises(ValueError):
        profiles.from_dict({"profile_version": profiles.PROFILE_VERSION + 1})


def test_documented_example_profile_is_valid():
    from pathlib import Path
    text = (Path(__file__).parents[1] / "config" / "example_profile.json").read_text()
    p = profiles.from_dict(json.loads(text))
    assert p.sheet_overrides["SHEET2"].amount.letter == "C"
    assert p.write.zero_fill_returns is False

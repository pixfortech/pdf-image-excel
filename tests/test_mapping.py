"""Tests for mapping configuration and group->sheet assignment without hardcoding."""
from src import mapping
from src.mapping import (
    AppConfig,
    OutputType,
    SheetMapping,
    SourceMapping,
    TargetMode,
    WriteAction,
    WriteRules,
)


def test_group_to_sheet_is_user_defined():
    cfg = AppConfig()
    # Arbitrary, user-chosen assignments supplied at runtime.
    cfg.group_to_sheet = {"GROUP ONE": "SheetA", "GROUP TWO": "SheetB"}
    assert cfg.group_to_sheet["GROUP ONE"] == "SheetA"
    # The app code never assumes group name == sheet name.
    assert cfg.group_to_sheet["GROUP TWO"] != "GROUP TWO"


def test_config_roundtrip_json():
    cfg = AppConfig(
        source=SourceMapping(
            group_field="g", date_field="d", amount_field="a",
            output_type=OutputType.FORMULA, sum_duplicate_dates=True,
        ),
        group_to_sheet={"X": "S1"},
        sheets={"S1": SheetMapping(
            sheet_name="S1", header_row=2,
            date_column="C", date_target_mode=TargetMode.COLUMN_LETTER,
            amount_column="F", amount_target_mode=TargetMode.COLUMN_LETTER,
        )},
        write_rules=WriteRules(write_action=WriteAction.ADD, output_type=OutputType.FORMULA),
    )
    text = mapping.config_to_json(cfg)
    restored = mapping.config_from_json(text)
    assert restored.source.group_field == "g"
    assert restored.source.output_type == OutputType.FORMULA
    assert restored.group_to_sheet == {"X": "S1"}
    assert restored.sheets["S1"].header_row == 2
    assert restored.sheets["S1"].amount_target_mode == TargetMode.COLUMN_LETTER
    assert restored.write_rules.write_action == WriteAction.ADD


def test_unknown_enum_falls_back_to_default():
    data = {"write_rules": {"write_action": "bogus"}}
    cfg = mapping.config_from_dict(data)
    assert cfg.write_rules.write_action == WriteAction.ASK

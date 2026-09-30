"""Drive the real Streamlit app.

Covers the monthly flow the app is built for: set up once (customers ->
worksheets, one worksheet pattern), then on a fresh app start the saved
profile is restored automatically and the preview is Ready without any
mapping — and no JSON upload.
"""
from pathlib import Path

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

from src.profiles import ProfileStore
from tests.conftest import make_report_pdf, make_workbook

APP = str(Path(__file__).parents[1] / "app.py")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MAPPING = {"ALPHA ONE": "S1", "BETA TWO": "S2", "GAMMA THREE": "S3"}


@pytest.fixture
def profile_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("PDFX_PROFILE_DIR", str(tmp_path / "profiles"))
    return tmp_path / "profiles"


def start_app():
    """A fresh app process: nothing cached, only the profile store on disk."""
    st.cache_resource.clear()
    return AppTest.from_file(APP, default_timeout=120).run()


def upload(at, pdf, xlsx):
    at.file_uploader[0].upload("report.pdf", pdf, "application/pdf")
    at.file_uploader[1].upload("book.xlsx", xlsx, XLSX_MIME)
    return at.run()


def texts(at):
    return " ".join(str(e.value) for kind in ("success", "warning", "error", "info") for e in getattr(at, kind))


def click(at, label):
    next(b for b in at.button if b.label == label).click()
    return at.run()


def status_table(at):
    return at.dataframe[0].value.to_string()


def set_up_once(at):
    for group, sheet in MAPPING.items():
        at.selectbox(key=f"fix_cust_{group}").set_value(sheet)
    at = click(at, "Save customer mapping")
    assert not at.exception
    # Pattern: taken from a mapped branch sheet; DATE and RETURN detected.
    assert at.selectbox(key="fix_pattern_sheet").value == "S1"
    assert at.selectbox(key="fix_pattern_S1_1_date").value == "A"
    assert at.selectbox(key="fix_pattern_S1_1_returns").value == "C"
    assert at.selectbox(key="fix_pattern_S1_1_amount").value is None      # never guessed
    at.selectbox(key="fix_pattern_S1_1_amount").set_value("B")
    return click(at, "Save worksheet pattern")


@pytest.fixture
def files(sample_groups):
    return make_report_pdf(sample_groups), make_workbook(["S1", "S2", "S3"], other_layout=["K1"])


def test_upload_prompt_without_files(profile_dir):
    at = start_app()
    assert not at.exception and "Upload the report and the workbook" in texts(at)


def test_set_up_once_then_restart_restores_everything_without_json(profile_dir, files):
    pdf, xlsx = files
    at = upload(start_app(), pdf, xlsx)
    assert "No saved mapping for this report yet" in texts(at)
    at = set_up_once(at)
    assert not at.exception
    assert "Needs attention" not in " ".join(m.value for m in at.markdown)
    assert "Not mapped" not in status_table(at)

    # --- restart the app; upload the same kind of files; no JSON import ---
    at = upload(start_app(), pdf, xlsx)
    assert not at.exception
    assert "Previous mapping recognised and loaded" in texts(at)
    table = status_table(at)
    assert "Not mapped" not in table and "Column problem" not in table
    assert table.count("Ready") == 3
    # The normal flow asks nothing: no mapping widgets outside "Advanced mapping".
    assert not [w for w in at.selectbox if (w.key or "").startswith("fix_")]
    assert "Needs attention" not in " ".join(m.value for m in at.markdown)
    patterns = " ".join(m.value for m in at.markdown)
    assert "DATE → A / DATE" in patterns and "Total Amount → B / CHALLAN" in patterns

    next(c for c in at.checkbox if c.label.startswith("I reviewed the preview")).check()
    at = click(at.run(), "Update uploaded workbook")
    assert not at.exception
    assert "Updated" in texts(at) and "Verified" in texts(at)


def test_only_the_differently_laid_out_sheet_needs_attention(profile_dir, files):
    pdf, xlsx = files
    at = upload(start_app(), pdf, xlsx)
    for group, sheet in {**MAPPING, "GAMMA THREE": "K1"}.items():
        at.selectbox(key=f"fix_cust_{group}").set_value(sheet)
    at = click(at, "Save customer mapping")
    at.selectbox(key="fix_pattern_S1_1_amount").set_value("B")
    at = click(at, "Save worksheet pattern")
    table = status_table(at)
    assert table.count("Column problem") == 1 and table.count("Ready") == 2   # only K1

    # The exception editor pre-fills by header on K1 (DATE=A, RETURN=D) and
    # leaves the ambiguous amount target for the user.
    assert at.selectbox(key="fix_exc_K1_1_date").value == "A"
    assert at.selectbox(key="fix_exc_K1_1_returns").value == "D"
    assert at.selectbox(key="fix_exc_K1_1_amount").value is None
    at.selectbox(key="fix_exc_K1_1_amount").set_value("B")
    at = click(at, "Save for K1")
    assert not at.exception
    assert "Column problem" not in status_table(at)
    saved = ProfileStore(profile_dir).all()[0]
    assert saved.sheet_overrides["K1"].amount.header == "KARKHANA"
    assert saved.sheet_template.amount.header == "CHALLAN"


def test_column_dropdown_follows_the_selected_worksheet(profile_dir, files):
    pdf, xlsx = files
    at = upload(start_app(), pdf, xlsx)
    for group, sheet in {**MAPPING, "GAMMA THREE": "K1"}.items():
        at.selectbox(key=f"fix_cust_{group}").set_value(sheet)
    at = click(at, "Save customer mapping")

    at.selectbox(key="adv_pattern_sheet").set_value("S1").run()
    s1_options = at.selectbox(key="adv_pattern_S1_1_amount").options
    assert any("CHALLAN" in o for o in s1_options)

    at.selectbox(key="adv_pattern_sheet").set_value("K1").run()
    k1_options = at.selectbox(key="adv_pattern_K1_1_amount").options
    assert any("KARKHANA" in o for o in k1_options)
    assert not any("CHALLAN" in o for o in k1_options)                    # nothing stale
    assert not [w for w in at.selectbox if w.key == "adv_pattern_S1_1_amount"]

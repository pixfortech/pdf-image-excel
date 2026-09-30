"""Drive the real Streamlit app: upload -> map once -> recognised -> update."""
from pathlib import Path

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

from tests.conftest import make_report_pdf, make_workbook

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@pytest.fixture
def app(tmp_path, monkeypatch):
    monkeypatch.setenv("PDFX_PROFILE_DIR", str(tmp_path / "profiles"))
    st.cache_resource.clear()
    at = AppTest.from_file(str(Path(__file__).parents[1] / "app.py"), default_timeout=120)
    return at.run()


def _texts(at):
    return " ".join(str(e.value) for kind in ("success", "warning", "error", "info")
                    for e in getattr(at, kind))


def _upload(at, sample_groups):
    at.file_uploader[0].upload("report.pdf", make_report_pdf(sample_groups), "application/pdf")
    at.file_uploader[1].upload("book.xlsx", make_workbook(["S1", "S2", "S3"]), XLSX_MIME)
    return at.run()


def test_differently_laid_out_sheet_is_resolved_once_in_the_ui(app, sample_groups, tmp_path):
    from src.profiles import ProfileStore
    from tests.conftest import BRANCH_LAYOUT
    app.file_uploader[0].upload("report.pdf", make_report_pdf(sample_groups), "application/pdf")
    app.file_uploader[1].upload("book.xlsx", make_workbook(["S1", "S2"], other_layout=["K1"]), XLSX_MIME)
    at = app.run()
    for group, sheet in {"ALPHA ONE": "S1", "BETA TWO": "S2", "GAMMA THREE": "K1"}.items():
        at.selectbox(key=f"grp_{group}").set_value(sheet)
    at.selectbox(key="tmpl_sheet").set_value("S1")
    for key, letter in (("tmpl_date", "A"), ("tmpl_amt", "B"), ("tmpl_ret", "C")):
        at.selectbox(key=key).set_value(letter)
    next(b for b in at.button if b.label == "Save mapping").click()
    at = at.run()
    assert "Column problem" in at.dataframe[0].value.to_string()     # only K1 is flagged

    # The override pre-fills by header (DATE=A, RETURN=D), and leaves the
    # ambiguous amount target empty for the user to decide.
    assert at.selectbox(key="ovr_K1_date").value == "A"
    assert at.selectbox(key="ovr_K1_ret").value == "D"
    assert at.selectbox(key="ovr_K1_amt").value is None
    at.selectbox(key="ovr_K1_amt").set_value("B")
    next(b for b in at.button if b.label == "Save mapping").click()
    at = at.run()
    assert not at.exception
    assert "Column problem" not in at.dataframe[0].value.to_string()
    saved = ProfileStore(tmp_path / "profiles").all()[0]
    assert saved.sheet_overrides["K1"].amount.header == "KARKHANA"
    assert saved.sheet_template == BRANCH_LAYOUT


def test_upload_prompt_without_files(app):
    assert not app.exception
    assert "Upload the report and the workbook" in _texts(app)


def test_first_run_then_recognised_then_update(app, sample_groups):
    at = _upload(app, sample_groups)
    assert not at.exception
    assert "No saved mapping matches this report yet" in _texts(at)

    # One-time mapping: customers -> sheets, and the worksheet column template.
    for group, sheet in {"ALPHA ONE": "S1", "BETA TWO": "S2", "GAMMA THREE": "S3"}.items():
        at.selectbox(key=f"grp_{group}").set_value(sheet)
    at.selectbox(key="tmpl_sheet").set_value("S1")
    at.selectbox(key="tmpl_date").set_value("A")
    at.selectbox(key="tmpl_amt").set_value("B")
    at.selectbox(key="tmpl_ret").set_value("C")
    next(b for b in at.button if b.label == "Save mapping").click()
    at = at.run()
    assert not at.exception
    assert "Previous mapping recognised and loaded" in _texts(at)
    assert "Needs attention" not in " ".join(m.value for m in at.markdown)

    update = next(b for b in at.button if b.label == "Update uploaded workbook")
    assert update.disabled                                   # one confirmation required
    next(c for c in at.checkbox if c.label.startswith("I reviewed the preview")).check()
    at = at.run()
    next(b for b in at.button if b.label == "Update uploaded workbook").click()
    at = at.run()
    assert not at.exception
    assert "Updated" in _texts(at) and "Verified" in _texts(at)

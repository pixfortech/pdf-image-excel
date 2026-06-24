"""End-to-end test that runs the full validation harness on generated real
files (a real PDF via reportlab + pdfplumber, a real .xlsx via openpyxl).

Skipped automatically if pdfplumber/reportlab are not installed, so the rest of
the suite still runs in minimal environments.
"""
import importlib

import pytest

pytest.importorskip("pdfplumber")
pytest.importorskip("reportlab")

validation = importlib.import_module("scripts.real_file_validation")


def test_full_pipeline_validation_runs_and_passes():
    # main() raises SystemExit only on a failed checkpoint; a clean run returns
    # None. We assert it completes without raising.
    validation.main()

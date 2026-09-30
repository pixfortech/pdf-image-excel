"""Local-only integration test on YOUR real files (skipped unless configured).

    PDFX_REAL_PDF=local/report.pdf PDFX_REAL_XLSX=local/book.xlsx \
    PDFX_REAL_PROFILE=local/profile.json PDFX_REAL_EXPECT=local/expect.json \
    pytest tests/test_real_files.py -s

Keep the files, profile and expectations in the gitignored ``local/`` folder.
"""
import os
from pathlib import Path

import pytest

from scripts.verify_real_files import run

REAL = {k: os.environ.get(f"PDFX_REAL_{k}") for k in ("PDF", "XLSX", "PROFILE", "EXPECT")}
pytestmark = pytest.mark.skipif(not (REAL["PDF"] and REAL["XLSX"]),
                                reason="set PDFX_REAL_PDF and PDFX_REAL_XLSX to run on real files")


def test_real_report_updates_real_workbook(tmp_path):
    summary = run(Path(REAL["PDF"]), Path(REAL["XLSX"]), REAL["PROFILE"], REAL["EXPECT"],
                  exclude_blocked=os.environ.get("PDFX_REAL_EXCLUDE_BLOCKED") == "1",
                  out_dir=tmp_path)
    assert summary["reconciled"], summary["reconciliation"]
    assert not summary["hard_blockers"]
    assert summary["written"] > 0
    assert summary["backup_intact"]
    assert summary["problems"] == []

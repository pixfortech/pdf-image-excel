import datetime as dt

import pytest

from src import extractor, parser, plan
from tests.conftest import make_report_pdf


def _parse(groups, **kw):
    doc = extractor.extract(make_report_pdf(groups, **{k: v for k, v in kw.items() if k != "labels"}), "r.pdf")
    return parser.parse(doc, header_labels=kw.get("labels", ()))


def _invoices(report):
    return parser.invoices(report, date_field="Doc Date", amount_field="Total Amount",
                           returns_field="Returns", reference_field="Ref No")


@pytest.mark.parametrize("ruled", [True, False], ids=["ruled-header", "word-gaps"])
def test_structure_detected_without_configuration(sample_groups, ruled):
    report = _parse(sample_groups, ruled=ruled)
    assert report.group_label == "Group Name"
    assert report.groups == list(sample_groups)                  # read dynamically, in order
    assert report.columns == ["Ref No", "Doc Date", "Total Amount", "Returns"]
    assert report.column_kind("Doc Date") == "date"
    assert report.column_kind("Total Amount") == "amount"
    assert report.column_kind("Ref No") == "text"
    assert report.column_kind("Returns") == "empty"
    assert report.period == (dt.date(2026, 8, 1), dt.date(2026, 8, 31))


def test_group_continues_across_pages_and_noise_is_excluded(sample_groups):
    report = _parse(sample_groups, rows_per_page=20)
    invs, rejected = _invoices(report)
    alpha = [i for i in invs if i.group == "ALPHA ONE"]
    assert len(alpha) == len(sample_groups["ALPHA ONE"])
    assert {i.page for i in alpha} >= {1, 2}                     # continued past a page break
    assert len(invs) == sum(len(r) for r in sample_groups.values())
    # Page footers are rejected; title/period are preamble; totals are not data.
    assert all("NoOfPages" in r.text for r in rejected)
    assert {i.reason for i in report.ignored} == {"before the first group"}
    assert report.group_total_label == "Group Total"
    assert report.grand_total is not None


def test_references_kept_for_audit_never_used_as_dates(sample_groups):
    invs, _ = _invoices(_parse(sample_groups))
    assert all(i.reference.startswith("REF-") for i in invs)
    assert all(isinstance(i.date, dt.date) for i in invs)


def test_blank_returns_stay_blank_and_return_only_rows_use_returns_column():
    groups = {"DELTA": [("R-1", "03/08/2026", "500.00", None),
                        ("R-2", "04/08/2026", None, "120.00")]}      # a return-only row
    invs, _ = _invoices(_parse(groups))
    first, second = invs
    assert first.amount == 500.0 and first.returns is None       # blank is NOT zero
    assert second.amount is None and second.returns == 120.0     # placed by position


def test_duplicate_dates_are_summed_with_breakup(sample_groups):
    invs, _ = _invoices(_parse(sample_groups))
    totals = {(t.group, t.date): t for t in plan.aggregate(invs)}
    day = totals[("ALPHA ONE", dt.date(2026, 8, 3))]
    assert day.amount == 19750.0
    assert day.breakup == "12,500.00 + 7,250.00"
    assert day.formula == "=12500+7250"
    assert day.references == "REF-1001, REF-1002"
    assert totals[("ALPHA ONE", dt.date(2026, 8, 4))].amount == 7000.0
    assert day.returns is None


def test_reconciles_with_printed_totals(sample_groups):
    report = _parse(sample_groups)
    recon = parser.reconcile(report, _invoices(report)[0])
    assert recon.checked and recon.ok
    scopes = [l.scope for l in recon.lines]
    assert "All invoices" in scopes and "ALPHA ONE" in scopes


def test_detects_a_total_that_does_not_reconcile(sample_groups):
    report = _parse(sample_groups, group_totals={"BETA TWO": ["3,600.00"]}, grand_total="99.00")
    recon = parser.reconcile(report, _invoices(report)[0])
    assert not recon.ok
    failed = {l.scope for l in recon.failures}
    assert "BETA TWO" in failed and "All invoices" in failed


def test_profile_header_labels_regroup_words_without_rules(sample_groups):
    report = _parse(sample_groups, ruled=False, labels=["Ref No", "Doc Date", "Total Amount", "Returns"])
    assert report.columns == ["Ref No", "Doc Date", "Total Amount", "Returns"]

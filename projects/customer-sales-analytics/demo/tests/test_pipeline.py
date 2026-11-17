import csv
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pipeline


@pytest.fixture
def inputs(tmp_path):
    directory = tmp_path / "input"
    pipeline.generate_sample(directory)
    return directory


def test_validation_reports_errors_and_excludes_all_duplicate_records(inputs):
    valid, issues = pipeline.validate(pipeline.load_inputs(inputs))
    assert set(valid["customers"]["customer_id"]) == {"C001", "C002", "C003", "C004"}
    assert set(valid["sales"]["sale_id"]) == {"SALE001", "SALE002", "SALE003", "SALE004"}
    assert set(valid["transactions"]["transaction_id"]) == {
        "TXN001", "TXN002", "TXN003", "TXN004", "TXN005", "TXN006",
    }
    assert (issues["issue"] == "duplicate identifier").sum() == 6
    assert len(issues) == 20
    assert {"missing required value", "invalid status", "unmatched identifier",
            "referenced record invalid", "invalid amount, precision, or sign"}.issubset(
        set(issues["issue"])
    )
    assert (issues["record_id"] == "SALE008").any()
    assert (issues["record_id"] == "TXN008").any()


def test_reconciliation_joins_and_fictional_totals(inputs):
    valid, _ = pipeline.validate(pipeline.load_inputs(inputs))
    sales, transactions, summary = pipeline.reconcile(valid)
    assert len(sales) == 4
    assert len(transactions) == 6
    assert sales.loc[sales["sale_id"] == "SALE003", "customer_status"].item() == "cancelled"
    by_rep = summary.set_index("representative_id").to_dict("index")
    assert by_rep["REP001"] == {
        "sales_count": 2, "transaction_count": 4, "gross_amount": "1800.00",
        "refund_amount": "-700.00", "net_amount": "1100.00", "fictional_rate": "0.05",
        "fictional_commission": "55.00",
    }
    assert by_rep["REP002"] == {
        "sales_count": 2, "transaction_count": 2, "gross_amount": "1250.00",
        "refund_amount": "0.00", "net_amount": "1250.00", "fictional_rate": "0.04",
        "fictional_commission": "50.00",
    }


def test_cancellation_does_not_invent_refund(inputs):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = frames["transactions"].loc[
        frames["transactions"]["transaction_id"] != "TXN005"
    ]
    valid, _ = pipeline.validate(frames)
    _, _, summary = pipeline.reconcile(valid)
    rep = summary.set_index("representative_id").loc["REP001"]
    assert (rep["gross_amount"], rep["refund_amount"], rep["net_amount"]) == (
        "1800.00", "-100.00", "1700.00"
    )


def test_duplicate_transaction_and_invalid_sale_are_never_joined(inputs):
    frames = pipeline.load_inputs(inputs)
    valid, issues = pipeline.validate(frames)
    _, transactions, summary = pipeline.reconcile(valid)
    assert "TXN012" not in set(transactions["transaction_id"])
    assert "SALE005" not in set(transactions["sale_id"])
    assert len(issues.loc[issues["record_id"] == "TXN012"]) == 2
    assert summary["transaction_count"].sum() == len(transactions)


def test_half_up_commission_and_amount_precision(inputs):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = pd.concat([
        frames["transactions"],
        pd.DataFrame([{"transaction_id": "TXN020", "sale_id": "SALE001",
                       "transaction_type": "sale", "amount": "0.10"},
                      {"transaction_id": "TXN021", "sale_id": "SALE001",
                       "transaction_type": "sale", "amount": "1.001"}]),
    ], ignore_index=True)
    valid, issues = pipeline.validate(frames)
    _, _, summary = pipeline.reconcile(valid)
    rep = summary.set_index("representative_id").loc["REP001"]
    assert rep["net_amount"] == "1100.10"
    assert rep["fictional_commission"] == "55.01"
    assert (issues["record_id"] == "TXN021").any()


def test_no_valid_sales_produces_empty_reports(inputs):
    frames = pipeline.load_inputs(inputs)
    frames["sales"]["customer_id"] = "UNKNOWN"
    valid, issues = pipeline.validate(frames)
    sales, transactions, summary = pipeline.reconcile(valid)
    assert sales.empty
    assert transactions.empty
    assert summary.empty
    assert not issues.empty


def test_missing_transaction_group_does_not_round_other_representative(inputs):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = pd.DataFrame([
        {"transaction_id": "TXN001", "sale_id": "SALE001",
         "transaction_type": "sale", "amount": "90071992547409.93"},
    ])
    valid, _ = pipeline.validate(frames)
    _, _, summary = pipeline.reconcile(valid)
    by_rep = summary.set_index("representative_id")
    assert by_rep.loc["REP001", "net_amount"] == "90071992547409.93"
    assert by_rep.loc["REP002", "net_amount"] == "0.00"


@pytest.mark.parametrize(
    ("kind", "amount", "field"),
    [
        ("sale", "50000000000000000.00", "gross_cents"),
        ("refund", "-50000000000000000.00", "refund_cents"),
    ],
)
def test_group_total_outside_integer_range_fails_explicitly(inputs, kind, amount, field):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = pd.DataFrame([
        {"transaction_id": "TXN001", "sale_id": "SALE001",
         "transaction_type": kind, "amount": amount},
        {"transaction_id": "TXN002", "sale_id": "SALE001",
         "transaction_type": kind, "amount": amount},
    ])
    valid, _ = pipeline.validate(frames)
    with pytest.raises(ValueError, match=field):
        pipeline.reconcile(valid)


@pytest.mark.parametrize("amount", ["92233720368547758.08", "-92233720368547758.08"])
def test_individual_amount_outside_integer_range_is_reported(inputs, amount):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = pd.DataFrame([
        {"transaction_id": "TXN001", "sale_id": "SALE001",
         "transaction_type": "sale" if not amount.startswith("-") else "refund",
         "amount": amount},
    ])
    valid, issues = pipeline.validate(frames)
    assert valid["transactions"].empty
    assert "amount exceeds supported range" in set(issues["issue"])


def test_maximum_supported_amount_is_reconciled_exactly(inputs):
    frames = pipeline.load_inputs(inputs)
    frames["transactions"] = pd.DataFrame([
        {"transaction_id": "TXN001", "sale_id": "SALE001",
         "transaction_type": "sale", "amount": "92233720368547758.07"},
    ])
    valid, _ = pipeline.validate(frames)
    _, _, summary = pipeline.reconcile(valid)
    assert summary.set_index("representative_id").loc["REP001", "net_amount"] == "92233720368547758.07"


def test_invalid_or_missing_schema_and_missing_file_are_explicit(inputs):
    (inputs / "customers.csv").write_text("customer_id\nC001\n", encoding="utf-8")
    with pytest.raises(ValueError, match="missing required columns"):
        pipeline.load_inputs(inputs)
    (inputs / "customers.csv").write_text("customer_id,customer_id\nC001,C001\n", encoding="utf-8")
    with pytest.raises(ValueError, match="duplicate column names"):
        pipeline.load_inputs(inputs)
    (inputs / "customers.csv").unlink()
    with pytest.raises(FileNotFoundError):
        pipeline.load_inputs(inputs)


def test_malformed_csv_is_not_silently_truncated(inputs):
    (inputs / "customers.csv").write_text(
        "customer_id,customer_status\nC001,active,extra\n", encoding="utf-8"
    )
    with pytest.raises(ValueError, match="invalid column count"):
        pipeline.load_inputs(inputs)


def test_reports_are_reproducible_and_workbook_is_readable(tmp_path):
    first = tmp_path / "first"
    second = tmp_path / "second"
    summary, issues = pipeline.run(first)
    second_summary, second_issues = pipeline.run(second)
    pd.testing.assert_frame_equal(summary, second_summary)
    pd.testing.assert_frame_equal(issues, second_issues)
    for path in first.rglob("*.csv"):
        assert path.read_bytes() == (second / path.relative_to(first)).read_bytes()
    published_example = Path(__file__).resolve().parents[1] / "examples" / "fictional_reconciliation_summary.csv"
    with published_example.open(newline="", encoding="utf-8") as expected:
        with (first / "reconciliation_summary.csv").open(newline="", encoding="utf-8") as actual:
            assert list(csv.reader(expected)) == list(csv.reader(actual))
    workbook_a = load_workbook(first / "fictional_reconciliation.xlsx", data_only=True)
    workbook_b = load_workbook(second / "fictional_reconciliation.xlsx", data_only=True)
    assert workbook_a.sheetnames == [
        "Read Me", "Validated Sales", "Valid Transactions", "Reconciliation", "Exceptions"
    ]
    for sheet_name in workbook_a.sheetnames:
        assert list(workbook_a[sheet_name].values) == list(workbook_b[sheet_name].values)
    assert len(workbook_a["Reconciliation"]._charts) == 1
    assert workbook_a["Reconciliation"]["D2"].value == 1800
    assert workbook_a["Reconciliation"]["E2"].value == -700
    assert workbook_a["Reconciliation"]["H2"].value == 55
    with (first / "exceptions.csv").open(newline="", encoding="utf-8") as stream:
        assert len(list(csv.DictReader(stream))) == len(issues)


def test_existing_output_directory_is_not_overwritten(tmp_path):
    output = tmp_path / "existing"
    output.mkdir()
    marker = output / "keep.txt"
    marker.write_text("keep", encoding="utf-8")
    with pytest.raises(FileExistsError):
        pipeline.run(output)
    assert marker.read_text(encoding="utf-8") == "keep"

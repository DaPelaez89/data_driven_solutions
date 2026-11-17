"""Generate and reconcile entirely fictional customer and sales data."""

import argparse
import csv
import sys
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


SCHEMAS = {
    "customers": ["customer_id", "customer_status"],
    "representatives": ["representative_id", "fictional_rate"],
    "sales": ["sale_id", "customer_id", "representative_id", "sale_status"],
    "transactions": ["transaction_id", "sale_id", "transaction_type", "amount"],
}
KEYS = {
    "customers": "customer_id",
    "representatives": "representative_id",
    "sales": "sale_id",
    "transactions": "transaction_id",
}
STATUSES = {"active", "pending", "cancelled"}
TYPES = {"sale", "refund"}
ISSUE_COLUMNS = ["entity", "source_row", "record_id", "field", "issue"]
MAX_CENTS = 2**63 - 1
MAX_AMOUNT = Decimal(MAX_CENTS) / 100

# Fixed invented records; none are taken from the existing case-study data.
SAMPLE = {
    "customers": [
        ["C001", " Active "], ["C002", "ACTIVE"], ["C003", "cancelled"],
        ["C004", "Pending"], ["", "active"], ["C006", "unknown"],
        ["C007", "active"], ["C007", "active"],
    ],
    "representatives": [["REP001", "0.05"], ["REP002", "0.04"]],
    "sales": [
        ["SALE001", "C001", "REP001", "active"],
        ["SALE002", "C002", "REP002", "ACTIVE"],
        ["SALE003", "C003", "REP001", "cancelled"],
        ["SALE004", "C004", "REP002", " active "],
        ["SALE005", "C999", "REP001", "active"],
        ["SALE006", "C001", "REP999", "active"],
        ["", "C001", "REP001", "active"],
        ["SALE008", "C007", "REP001", "active"],
        ["SALE009", "C006", "REP001", "active"],
        ["SALE010", "C001", "REP001", "unknown"],
        ["SALE011", "C001", "REP001", "active"],
        ["SALE011", "C001", "REP001", "active"],
    ],
    "transactions": [
        ["TXN001", "SALE001", "sale", "1200.00"],
        ["TXN002", "SALE001", "refund", "-100.00"],
        ["TXN003", "SALE002", "sale", "850.00"],
        ["TXN004", "SALE003", "sale", "600.00"],
        ["TXN005", "SALE003", "refund", "-600.00"],
        ["TXN006", "SALE004", "sale", "400.00"],
        ["TXN007", "SALE999", "sale", "200.00"],
        ["TXN008", "SALE005", "sale", "300.00"],
        ["TXN009", "SALE001", "sale", "invalid"],
        ["", "SALE001", "sale", "50.00"],
        ["TXN011", "SALE002", "refund", "25.00"],
        ["TXN012", "SALE001", "sale", "999.00"],
        ["TXN012", "SALE001", "sale", "999.00"],
        ["TXN013", "SALE002", "other", "30.00"],
    ],
}


def generate_sample(input_dir: Path) -> None:
    """Write the fixed fictional input files to a new directory."""
    input_dir.mkdir(parents=True, exist_ok=False)
    for entity, columns in SCHEMAS.items():
        with (input_dir / f"{entity}.csv").open("w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream, lineterminator="\n")
            writer.writerow(columns)
            writer.writerows(SAMPLE[entity])


def load_inputs(input_dir: Path) -> dict[str, pd.DataFrame]:
    """Load strings while rejecting missing, duplicate, or malformed headers."""
    frames = {}
    for entity, columns in SCHEMAS.items():
        path = input_dir / f"{entity}.csv"
        with path.open(encoding="utf-8-sig", newline="") as stream:
            reader = csv.reader(stream, strict=True)
            header = next(reader, [])
            if not header or any(not name for name in header) or len(header) != len(set(header)):
                raise ValueError(f"{path.name}: missing, empty, or duplicate column names")
            missing = sorted(set(columns) - set(header))
            if missing:
                raise ValueError(f"{path.name}: missing required columns: {', '.join(missing)}")
            for record_number, row in enumerate(reader, start=2):
                if len(row) != len(header):
                    raise ValueError(f"{path.name}: invalid column count in record {record_number}")
        frames[entity] = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    return frames


def _decimal(value: str) -> Decimal:
    try:
        result = Decimal(value)
    except InvalidOperation as error:
        raise ValueError("not a decimal") from error
    if not result.is_finite():
        raise ValueError("not finite")
    return result


def validate(frames: dict[str, pd.DataFrame]) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    """Normalize strings, report exceptions, and exclude all ambiguous records."""
    normalized = {name: frame.copy() for name, frame in frames.items()}
    errors = {name: set() for name in SCHEMAS}
    issues = []

    def reject(entity: str, index: int, field: str, issue: str) -> None:
        errors[entity].add(index)
        issues.append({
            "entity": entity,
            "source_row": index + 2,
            "record_id": normalized[entity].at[index, KEYS[entity]],
            "field": field,
            "issue": issue,
        })

    for entity, columns in SCHEMAS.items():
        frame = normalized[entity]
        for column in columns:
            frame[column] = frame[column].str.strip()
            if column.endswith("_id"):
                frame[column] = frame[column].str.upper()
            elif column.endswith("_status") or column == "transaction_type":
                frame[column] = frame[column].str.lower()
            for index in frame.index[frame[column] == ""]:
                reject(entity, index, column, "missing required value")
        key = KEYS[entity]
        duplicate = frame[key].ne("") & frame[key].duplicated(keep=False)
        for index in frame.index[duplicate]:
            reject(entity, index, key, "duplicate identifier")

    for entity, field in [("customers", "customer_status"), ("sales", "sale_status")]:
        frame = normalized[entity]
        for index in frame.index[frame[field].ne("") & ~frame[field].isin(STATUSES)]:
            reject(entity, index, field, "invalid status")

    reps = normalized["representatives"]
    for index, value in reps["fictional_rate"].items():
        if not value:
            continue
        try:
            rate = _decimal(value)
            if not Decimal("0") <= rate <= Decimal("1"):
                raise ValueError("out of range")
        except ValueError:
            reject("representatives", index, "fictional_rate", "invalid rate (expected 0 to 1)")

    def valid_ids(entity: str) -> set[str]:
        frame = normalized[entity]
        return set(frame.loc[~frame.index.isin(errors[entity]), KEYS[entity]])

    sales = normalized["sales"]
    for field, referenced in [("customer_id", "customers"), ("representative_id", "representatives")]:
        allowed = valid_ids(referenced)
        all_ids = set(normalized[referenced][KEYS[referenced]])
        for index, value in sales[field].items():
            if value and value not in allowed:
                issue = "referenced record invalid" if value in all_ids else "unmatched identifier"
                reject("sales", index, field, issue)

    transactions = normalized["transactions"]
    allowed_sales = valid_ids("sales")
    all_sales = set(sales["sale_id"])
    for index, row in transactions.iterrows():
        sale_id = row["sale_id"]
        if sale_id and sale_id not in allowed_sales:
            issue = "referenced record invalid" if sale_id in all_sales else "unmatched identifier"
            reject("transactions", index, "sale_id", issue)
        kind = row["transaction_type"]
        if kind and kind not in TYPES:
            reject("transactions", index, "transaction_type", "invalid transaction type")
        amount = row["amount"]
        if amount:
            try:
                number = _decimal(amount)
                if number.as_tuple().exponent < -2:
                    raise ValueError("more than two decimal places")
                if kind in TYPES and ((kind == "sale" and number <= 0) or (kind == "refund" and number >= 0)):
                    raise ValueError("incorrect sign")
            except ValueError:
                reject("transactions", index, "amount", "invalid amount, precision, or sign")
            else:
                if number.copy_abs() > MAX_AMOUNT:
                    reject("transactions", index, "amount", "amount exceeds supported range")

    valid = {
        entity: frame.loc[~frame.index.isin(errors[entity])].copy()
        for entity, frame in normalized.items()
    }
    return valid, pd.DataFrame(issues, columns=ISSUE_COLUMNS)


def reconcile(valid: dict[str, pd.DataFrame]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Join validated data and compute fictional amounts using integer cents."""
    sales = valid["sales"].merge(
        valid["customers"], on="customer_id", how="left", validate="many_to_one"
    ).merge(
        valid["representatives"], on="representative_id", how="left", validate="many_to_one"
    )
    transactions = valid["transactions"].merge(
        sales, on="sale_id", how="left", validate="many_to_one"
    )
    transactions["amount_cents"] = transactions["amount"].map(lambda value: int(_decimal(value) * 100))
    transactions["gross_cents"] = transactions["amount_cents"].where(
        transactions["transaction_type"] == "sale", 0
    )
    transactions["refund_cents"] = transactions["amount_cents"].where(
        transactions["transaction_type"] == "refund", 0
    )
    grouped = transactions.groupby("representative_id", as_index=False)
    for representative, group in grouped:
        for column in ["gross_cents", "refund_cents"]:
            total = sum(int(value) for value in group[column])
            if not -MAX_CENTS <= total <= MAX_CENTS:
                raise ValueError(f"{column} exceeds supported range for representative {representative}")
    totals = grouped.agg(
        transaction_count=("transaction_id", "size"),
        gross_cents=("gross_cents", "sum"),
        refund_cents=("refund_cents", "sum"),
    )
    for column in ["transaction_count", "gross_cents", "refund_cents"]:
        totals[column] = totals[column].astype("Int64")
    counts = sales.groupby("representative_id", as_index=False).agg(sales_count=("sale_id", "size"))
    summary = counts.merge(totals, on="representative_id", how="left", validate="one_to_one")
    summary = summary.merge(
        valid["representatives"], on="representative_id", how="left", validate="one_to_one"
    ).sort_values("representative_id").reset_index(drop=True)
    for column in ["transaction_count", "gross_cents", "refund_cents"]:
        summary[column] = summary[column].fillna(0).astype("int64")
    summary["net_cents"] = summary["gross_cents"] + summary["refund_cents"]

    def money(cents: int) -> str:
        return f"{Decimal(int(cents)) / 100:.2f}"

    for name in ["gross", "refund", "net"]:
        summary[f"{name}_amount"] = summary[f"{name}_cents"].map(money)
    summary["fictional_commission"] = [
        str((Decimal(int(cents)) / 100 * _decimal(rate)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
        for cents, rate in zip(summary["net_cents"], summary["fictional_rate"])
    ]
    summary = summary[[
        "representative_id", "sales_count", "transaction_count", "gross_amount",
        "refund_amount", "net_amount", "fictional_rate", "fictional_commission",
    ]]
    transactions = transactions.drop(columns=["amount_cents", "gross_cents", "refund_cents"])
    return sales, transactions, summary


def write_reports(output_dir: Path, sales: pd.DataFrame, transactions: pd.DataFrame,
                  summary: pd.DataFrame, issues: pd.DataFrame) -> None:
    """Write synthetic CSV reports and a workbook with a fictional-data chart."""
    reports = {
        "validated_sales": sales,
        "validated_transactions": transactions,
        "reconciliation_summary": summary,
        "exceptions": issues,
    }
    for name, frame in reports.items():
        frame.to_csv(output_dir / f"{name}.csv", index=False, lineterminator="\n")

    notes = pd.DataFrame({"Information": [
        "All records, amounts, rates, and results are fictional.",
        "Refunds are negative transactions; cancellation alone does not create a refund.",
        "Excluded records appear in the Exceptions sheet and are not included in totals.",
        "Commission is an invented rate multiplied by net transactions, rounded half up.",
    ]})
    workbook_summary = summary.copy()
    for column in ["gross_amount", "refund_amount", "net_amount", "fictional_commission"]:
        workbook_summary[column] = workbook_summary[column].map(float)
    with pd.ExcelWriter(output_dir / "fictional_reconciliation.xlsx", engine="openpyxl") as writer:
        notes.to_excel(writer, sheet_name="Read Me", index=False)
        sales.to_excel(writer, sheet_name="Validated Sales", index=False)
        transactions.to_excel(writer, sheet_name="Valid Transactions", index=False)
        workbook_summary.to_excel(writer, sheet_name="Reconciliation", index=False)
        issues.to_excel(writer, sheet_name="Exceptions", index=False)
        sheet = writer.sheets["Reconciliation"]
        for row in sheet.iter_rows(min_row=2):
            for position in [4, 5, 6, 8]:
                row[position - 1].number_format = "0.00"
        if not summary.empty:
            chart = BarChart()
            chart.title = "Fictional sales, refunds and net"
            chart.y_axis.title = "Fictional amount"
            chart.add_data(Reference(sheet, min_col=4, max_col=6, min_row=1, max_row=len(summary) + 1), titles_from_data=True)
            chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=len(summary) + 1))
            sheet.add_chart(chart, "J2")
        for worksheet in writer.sheets.values():
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for column in worksheet.columns:
                longest = max(len(str(cell.value or "")) for cell in column)
                worksheet.column_dimensions[get_column_letter(column[0].column)].width = min(longest + 2, 55)


def run(output_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Generate inputs and reports without overwriting an existing directory."""
    output_dir.mkdir(parents=True, exist_ok=False)
    input_dir = output_dir / "synthetic_inputs"
    generate_sample(input_dir)
    valid, issues = validate(load_inputs(input_dir))
    sales, transactions, summary = reconcile(valid)
    write_reports(output_dir, sales, transactions, summary, issues)
    return summary, issues


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--output-dir", type=Path, default=Path(__file__).resolve().parent / "generated",
        help="New directory for fictional inputs and reports (must not already exist)",
    )
    args = parser.parse_args()
    try:
        summary, issues = run(args.output_dir)
    except (OSError, ValueError, csv.Error, pd.errors.ParserError) as error:
        parser.exit(1, f"Pipeline failed: {error}\n")
    print(summary.to_string(index=False))
    print(f"\nFictional exceptions: {len(issues)}; reports: {args.output_dir}")


if __name__ == "__main__":
    main()

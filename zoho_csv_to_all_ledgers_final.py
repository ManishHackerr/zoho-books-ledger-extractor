"""
Zoho Books CSV to plain All Ledgers workbook.
Put this script and the Zoho Detailed General Ledger CSV in the same folder.
Output: Zoho_All_Ledgers_01Apr2026_to_30Sep2026.xlsx
"""

from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Change these only if you need a different period or organization name.
START_DATE = "2026-04-01"
END_DATE = "2026-09-30"
COMPANY_NAME = "Zoho Books"
OUTPUT_FILENAME = "Zoho_All_Ledgers_01Apr2026_to_30Sep2026.xlsx"


def locate_csv(folder):
    matches = sorted(folder.glob("*Detailed General Ledger*.csv"))
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise RuntimeError(
            "More than one Detailed General Ledger CSV found. "
            "Keep only the CSV you want to process in this folder."
        )
    files = sorted(folder.glob("*.csv"))
    if len(files) == 1:
        return files[0]
    if not files:
        raise FileNotFoundError(
            "No CSV found. Put the exported Detailed General Ledger CSV "
            "in the same folder as this script."
        )
    raise RuntimeError(
        "Multiple CSV files found. Keep only the required export in this folder."
    )


def text_value(row, column):
    if column not in row.index:
        return None
    value = row[column]
    if pd.isna(value) or not str(value).strip():
        return None
    return str(value).strip()


def main():
    folder = Path(__file__).resolve().parent
    csv_path = locate_csv(folder)
    output_path = folder / OUTPUT_FILENAME

    df = pd.read_csv(csv_path, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip().lower() for c in df.columns]

    required = {"date", "account_name", "debit", "credit"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(
            "CSV columns missing: " + ", ".join(sorted(missing))
        )

    df["parsed_date"] = pd.to_datetime(df["date"], errors="coerce", dayfirst=False)
    df = df[df["parsed_date"].between(
        pd.Timestamp(START_DATE), pd.Timestamp(END_DATE), inclusive="both"
    )].copy()

    if df.empty:
        raise ValueError(
            f"No transactions found between {START_DATE} and {END_DATE}. "
            "Check the CSV date format and date settings."
        )

    df["debit_num"] = pd.to_numeric(
        df["debit"].replace("", pd.NA), errors="coerce"
    ).fillna(0)
    df["credit_num"] = pd.to_numeric(
        df["credit"].replace("", pd.NA), errors="coerce"
    ).fillna(0)

    sort_cols = ["account_name", "parsed_date"]
    if "transaction_id" in df.columns:
        sort_cols.append("transaction_id")
    df = df.sort_values(sort_cols, kind="stable")

    wb = Workbook()
    ws = wb.active
    ws.title = "All Ledgers"
    ws.sheet_view.showGridLines = False

    ws.append([COMPANY_NAME])
    ws.append([
        "Ledger Report - "
        f"{pd.Timestamp(START_DATE).strftime('%d-%b-%Y')} to "
        f"{pd.Timestamp(END_DATE).strftime('%d-%b-%Y')}"
    ])
    ws.append([])

    headers = [
        "Date", "By/To", "Particulars", "Vch Type",
        "Vch No.", "Debit", "Credit", "Narration"
    ]
    period_label = (
        f"{pd.Timestamp(START_DATE).day}-{pd.Timestamp(START_DATE).strftime('%b-%y')} "
        f"to {pd.Timestamp(END_DATE).day}-{pd.Timestamp(END_DATE).strftime('%b-%y')}"
    )

    for account_name, entries in df.groupby(
        "account_name", sort=True, dropna=False
    ):
        ledger_name = str(account_name).strip() or "(Blank account name)"
        ws.append([ledger_name])
        ws.append([period_label])
        ws.append(headers)

        for _, row in entries.iterrows():
            debit = float(row["debit_num"])
            credit = float(row["credit_num"])
            by_to = (
                "By" if credit != 0 and debit == 0
                else "To" if debit != 0 and credit == 0
                else ""
            )

            particulars = (
                text_value(row, "transaction_details")
                or text_value(row, "description")
                or text_value(row, "reference_number")
                or text_value(row, "entity_number")
            )

            voucher_type = text_value(row, "transaction_type")
            if voucher_type:
                voucher_type = voucher_type.replace("_", " ").title()

            voucher_number = (
                text_value(row, "reference_number")
                or text_value(row, "entity_number")
            )

            ws.append([
                row["parsed_date"].strftime("%d-%m-%Y"),
                by_to,
                particulars,
                voucher_type,
                voucher_number,
                debit if debit != 0 else None,
                credit if credit != 0 else None,
                text_value(row, "description"),
            ])

        # Blank line between ledgers; no totals or decorative rows.
        ws.append([])

    # Plain appearance only. Assign fresh styles; do not mutate immutable style objects.
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Arial", size=10, bold=False)
            if cell.column in (6, 7) and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00;[Red](#,##0.00);-'

    ws["A1"].font = Font(name="Arial", size=10, bold=True)

    # Bold ledger names only; header labels remain plain, matching a simple ledger export.
    for row_num in range(4, ws.max_row + 1):
        if (
            ws.cell(row_num, 1).value
            and all(ws.cell(row_num, col).value is None for col in range(2, 9))
        ):
            ws.cell(row_num, 1).font = Font(name="Arial", size=10, bold=True)

    widths = {1: 14, 2: 10, 3: 38, 4: 16, 5: 22, 6: 15, 7: 15, 8: 55}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width

    wb.save(output_path)

    print("\nSUCCESS")
    print(f"CSV: {csv_path.name}")
    print(f"Excel created: {output_path.name}")
    print(f"Transactions: {len(df):,}")
    print(f"Ledgers: {df['account_name'].nunique():,}")
    print(f"Period: {START_DATE} to {END_DATE}")
    print(
        "\nNote: This CSV does not contain account groups or opening/closing "
        "balances, so the script does not invent them."
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}")
    input("\nPress Enter to close this window...")

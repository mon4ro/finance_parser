from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from finance_parser.common import clean_for_excel


DEFAULT_INPUT = Path("output/budgeting/ParsedTransactions.xlsx")
DEFAULT_OUTPUT = Path("output/budgeting/ParsedTransactions_values_only_test.xlsx")


DATE_HEADER_NAMES = {
    "Date",
    "Value Date",
    "ValueDate",
    "Booking Date",
    "BookingDate",
    "ExportDate",
    "ImportedAt",
    "Parsing date",
}

AMOUNT_HEADER_NAMES = {
    "Amount",
    "Amount EUR",
    "Price",
    "Fees",
    "Total price",
    "Charge",
    "Rate",
}


def read_workbook_values_only(path: Path) -> dict[str, list[list[object]]]:
    """
    Read workbook sheet values only.

    The loaded workbook object is closed and never saved. This function is
    intentionally one-way: existing workbook package XML is not round-tripped.
    """
    wb = load_workbook(path, data_only=False, read_only=True)
    try:
        sheets: dict[str, list[list[object]]] = {}
        for ws in wb.worksheets:
            rows: list[list[object]] = []
            for row in ws.iter_rows(values_only=True):
                rows.append([cell for cell in row])
            sheets[ws.title] = rows
        return sheets
    finally:
        wb.close()


def _non_empty_width(value: object) -> int:
    if value is None:
        return 0
    return min(max(len(str(value)), 0), 60)


def apply_basic_formatting(ws) -> None:
    """
    Apply simple deterministic formatting generated from code.

    This intentionally avoids copying any formatting/styles/tables from the old
    workbook package.
    """
    if ws.max_row < 1 or ws.max_column < 1:
        return

    ws.freeze_panes = "A2"

    end_col = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{end_col}{ws.max_row}"

    for cell in ws[1]:
        cell.font = Font(bold=True)

    headers = [ws.cell(row=1, column=col_idx).value for col_idx in range(1, ws.max_column + 1)]

    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        sample_widths = [_non_empty_width(header)]
        for row_idx in range(2, min(ws.max_row, 100) + 1):
            sample_widths.append(_non_empty_width(ws.cell(row=row_idx, column=col_idx).value))

        width = min(max(max(sample_widths, default=0) + 2, 8), 45)
        ws.column_dimensions[col_letter].width = width

        header_text = "" if header is None else str(header)
        if header_text in DATE_HEADER_NAMES:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = "yyyy-mm-dd"
        elif header_text in AMOUNT_HEADER_NAMES:
            for row_idx in range(2, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_idx).number_format = '#,##0.00'


def safe_table_name(sheet_name: str, existing: set[str]) -> str:
    """
    Generate a valid-ish unique Excel table display name.

    Excel table names must start with a letter or underscore and cannot contain
    spaces or many punctuation characters. Keep this simple and deterministic.
    """
    base = re.sub(r"[^A-Za-z0-9_]", "_", sheet_name).strip("_")
    if not base:
        base = "Sheet"
    if not re.match(r"^[A-Za-z_]", base):
        base = "T_" + base

    name = "tbl_" + base
    name = name[:240]

    candidate = name
    counter = 1
    while candidate in existing:
        suffix = f"_{counter}"
        candidate = name[: 255 - len(suffix)] + suffix
        counter += 1

    existing.add(candidate)
    return candidate


def apply_excel_table(ws, *, table_name: str) -> bool:
    """
    Add a fresh generated Excel table to a worksheet.

    Returns True if a table was added. A table is skipped if the sheet does not
    have at least one header row and one data row.
    """
    if ws.max_row < 2 or ws.max_column < 1:
        return False

    seen_headers: set[str] = set()
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        header = "" if cell.value is None else str(cell.value).strip()
        if not header:
            header = f"Column{col_idx}"
        original = header
        n = 2
        while header in seen_headers:
            header = f"{original}_{n}"
            n += 1
        seen_headers.add(header)
        cell.value = clean_for_excel(header)

    end_col = get_column_letter(ws.max_column)
    ref = f"A1:{end_col}{ws.max_row}"

    table = Table(displayName=table_name, ref=ref)
    style = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    table.tableStyleInfo = style
    ws.add_table(table)
    return True


def write_fresh_value_workbook(
    path: Path,
    sheets: dict[str, list[list[object]]],
    *,
    basic_formatting: bool = False,
    excel_tables: bool = False,
) -> dict[str, int]:
    """
    Create a brand-new workbook from plain values.

    This never copies the old .xlsx file and never saves a workbook object loaded
    from the old file.
    """
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    stats = {
        "sheets_written": 0,
        "tables_added": 0,
    }
    existing_table_names: set[str] = set()

    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        stats["sheets_written"] += 1

        for row in rows:
            ws.append([clean_for_excel(value) for value in row])

        if basic_formatting:
            apply_basic_formatting(ws)
        elif rows and rows[0]:
            ws.freeze_panes = "A2"

        if excel_tables:
            table_name = safe_table_name(sheet_name, existing_table_names)
            if apply_excel_table(ws, table_name=table_name):
                stats["tables_added"] += 1

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    wb.close()
    return stats


def validate_xlsx_basic(path: Path, *, expected_sheets: set[str] | None = None) -> list[str]:
    problems: list[str] = []

    if not path.exists():
        return [f"File does not exist: {path}"]

    if not zipfile.is_zipfile(path):
        return [f"Not a valid zip/xlsx package: {path}"]

    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            if not name.endswith(".xml"):
                continue
            try:
                ET.fromstring(zf.read(name))
            except Exception as exc:
                problems.append(f"Invalid XML part {name}: {type(exc).__name__}: {exc}")

    try:
        wb = load_workbook(path, read_only=True, data_only=False)
        sheetnames = set(wb.sheetnames)
        wb.close()
    except Exception as exc:
        problems.append(f"openpyxl could not reopen workbook: {type(exc).__name__}: {exc}")
        return problems

    if expected_sheets:
        missing = sorted(expected_sheets - sheetnames)
        if missing:
            problems.append("Missing expected sheet(s): " + ", ".join(missing))

    return problems


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Create a fresh value-only workbook copy without round-tripping the "
            "existing xlsx package. Diagnostic utility for Excel repair/corruption issues."
        )
    )
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_INPUT,
        help=f"Existing workbook to read as values only. Default: {DEFAULT_INPUT}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT,
        help=f"Fresh value-only workbook to write. Default: {DEFAULT_OUTPUT}",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Allow overwriting the output file if it already exists.",
    )
    parser.add_argument(
        "--basic-formatting",
        action="store_true",
        help=(
            "Apply simple generated formatting: bold headers, freeze panes, "
            "autofilter, column widths, and basic date/number formats."
        ),
    )
    parser.add_argument(
        "--excel-tables",
        action="store_true",
        help=(
            "Create fresh generated Excel Table objects for sheets that have "
            "a header row and at least one data row. Does not copy old tables."
        ),
    )
    args = parser.parse_args()

    input_path = args.input.expanduser().resolve()
    output_path = args.output.expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    if output_path.exists() and not args.overwrite:
        raise FileExistsError(
            f"Output already exists: {output_path}. Use --overwrite to replace it."
        )

    print(f"Read previous workbook values only: {input_path}")
    sheets = read_workbook_values_only(input_path)
    print(f"Loaded sheets: {len(sheets)}")
    for name, rows in sheets.items():
        print(f"  {name}: {len(rows)} row(s)")

    print("No loaded workbook object will be saved.")
    print("Creating fresh workbook from scratch.")
    if args.basic_formatting:
        print("Applying basic formatting generated from code.")
    else:
        print("Writing value-only workbook without basic formatting.")

    if args.excel_tables:
        print("Adding fresh generated Excel Table objects.")

    stats = write_fresh_value_workbook(
        output_path,
        sheets,
        basic_formatting=args.basic_formatting,
        excel_tables=args.excel_tables,
    )
    print(f"Sheets written: {stats['sheets_written']}")
    print(f"Excel tables added: {stats['tables_added']}")
    print(f"Saved fresh workbook: {output_path}")

    expected = set(sheets.keys())
    problems = validate_xlsx_basic(output_path, expected_sheets=expected)
    if problems:
        print("Validation problems:")
        for problem in problems:
            print(f"  - {problem}")
        raise SystemExit(1)

    print("Validated fresh workbook: OK")
    print("")
    print("Next manual test:")
    print("Open this file in Excel and check whether Excel reports repair/corruption:")
    print(f"  {output_path}")


if __name__ == "__main__":
    main()

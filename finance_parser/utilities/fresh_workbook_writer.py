from __future__ import annotations

import re
import shutil
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

from finance_parser.common import clean_for_excel


REVIEW_STATUS_SHEET = "UnifiedTransactions"

REVIEW_STATUS_REQUIRED_HEADERS = [
    "Include",
    "Owner",
    "Supercategory",
    "Category",
    "Subcategory",
    "NormalizedReceiver",
    "Review/Notes",
    "ReviewStatus",
]


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
    Read all worksheet values from an existing workbook.

    The workbook object loaded from the existing file is closed and never saved.
    This is the core anti-corruption guarantee: existing workbook XML packages
    are treated as read-only data sources.
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


def sheet_values_to_records(rows: list[list[object]]) -> tuple[list[str], list[dict[str, object]]]:
    if not rows:
        return [], []

    headers = ["" if value is None else str(value).strip() for value in rows[0]]
    records: list[dict[str, object]] = []

    for row in rows[1:]:
        record: dict[str, object] = {}
        for idx, header in enumerate(headers):
            if not header:
                continue
            record[header] = row[idx] if idx < len(row) else None
        records.append(record)

    return headers, records


def records_to_sheet_values(headers: list[str], records: list[dict[str, object]]) -> list[list[object]]:
    rows: list[list[object]] = [list(headers)]
    for record in records:
        rows.append([record.get(header) for header in headers])
    return rows


def append_changelog_row(
    sheets: dict[str, list[list[object]]],
    *,
    script: str,
    action: str,
    sheet: str,
    rows_updated: int,
    status: str,
    details: str,
    backup_file: Path | None = None,
) -> None:
    """
    Append ChangeLog as values before writing a fresh workbook.

    This avoids reopening and saving the workbook after the main write.
    """
    default_headers = [
        "Timestamp",
        "Script",
        "Action",
        "Sheet",
        "RowsUpdated",
        "BackupFile",
        "Status",
        "Details",
    ]

    rows = sheets.get("ChangeLog")
    if not rows:
        rows = [default_headers]
        sheets["ChangeLog"] = rows
        headers = default_headers
    else:
        headers = ["" if value is None else str(value).strip() for value in rows[0]]
        if not any(headers):
            rows[0] = default_headers
            headers = default_headers

    values = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Script": script,
        "Action": action,
        "Sheet": sheet,
        "RowsUpdated": rows_updated,
        "BackupFile": "" if backup_file is None else str(backup_file),
        "Status": status,
        "Details": details,
    }
    rows.append([values.get(header, "") for header in headers])


def _non_empty_width(value: object) -> int:
    if value is None:
        return 0
    return min(max(len(str(value)), 0), 60)


def apply_basic_formatting(ws) -> None:
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


def _header_column_letters(ws) -> dict[str, str]:
    letters: dict[str, str] = {}
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        if header:
            letters[str(header).strip()] = get_column_letter(col_idx)
    return letters


def _review_status_conditional_formatting(ws, column_letter: str, max_row: int) -> None:
    """
    Apply Excel's built-in Good/Neutral/Bad traffic-light colours (light
    background, darker text) to the ReviewStatus column.
    """
    cell_range = f"{column_letter}2:{column_letter}{max_row}"

    good_fill = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
    good_font = Font(color="FF006100")
    neutral_fill = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")
    neutral_font = Font(color="FF9C6500")
    bad_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    bad_font = Font(color="FF9C0006")

    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"READY"'], fill=good_fill, font=good_font),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"SKIP"'], fill=neutral_fill, font=neutral_font),
    )
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="equal", formula=['"TODO"'], fill=bad_fill, font=bad_font),
    )


def apply_review_status_column(ws) -> bool:
    """
    Regenerate the ReviewStatus formula and its traffic-light conditional
    formatting for every UnifiedTransactions data row.

    The fresh-rebuild write path only ever carries raw cell values forward -
    any formula or conditional formatting a user adds directly in Excel is
    silently flattened/lost on the next pipeline run (the formula's cached
    result survives as dead text; the formula itself and any conditional
    formatting do not). ReviewStatus is therefore treated as a pipeline-owned
    derived column, like Month/Year, except the derivation is written as a
    live Excel formula rather than a static Python-computed value, so it
    keeps updating interactively while a user edits Owner/Category/etc.
    directly in Excel between pipeline runs.

    Column positions are looked up by header name, not hardcoded letters, so
    this stays correct if UNIFIED_COLUMNS is ever reordered.

    Returns True if the formula/formatting was applied, False if this sheet
    doesn't have the expected columns (e.g. called on the wrong sheet) or has
    no data rows.
    """
    if ws.max_row < 2:
        return False

    letters = _header_column_letters(ws)
    if not set(REVIEW_STATUS_REQUIRED_HEADERS).issubset(letters):
        return False

    include_c = letters["Include"]
    owner_c = letters["Owner"]
    super_c = letters["Supercategory"]
    category_c = letters["Category"]
    sub_c = letters["Subcategory"]
    norm_c = letters["NormalizedReceiver"]
    notes_c = letters["Review/Notes"]
    status_c = letters["ReviewStatus"]

    for row_idx in range(2, ws.max_row + 1):
        formula = (
            f'=IF({include_c}{row_idx}="NO","SKIP",'
            f'IF(AND({include_c}{row_idx}="YES",{owner_c}{row_idx}<>"",{super_c}{row_idx}<>"",'
            f'{category_c}{row_idx}<>"",{sub_c}{row_idx}<>"",{norm_c}{row_idx}<>"",{notes_c}{row_idx}<>""),'
            f'"READY","TODO"))'
        )
        ws[f"{status_c}{row_idx}"] = formula

    _review_status_conditional_formatting(ws, status_c, ws.max_row)
    return True


def safe_table_name(sheet_name: str, existing: set[str]) -> str:
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
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)
    return True


def write_fresh_workbook(
    path: Path,
    sheets: dict[str, list[list[object]]],
    *,
    basic_formatting: bool = True,
    excel_tables: bool = True,
) -> dict[str, int]:
    """
    Create a brand-new workbook from sheet values.

    This never copies the old .xlsx package and never saves a workbook object
    loaded from the old file.
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

        if sheet_name == REVIEW_STATUS_SHEET:
            apply_review_status_column(ws)

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


def timestamped_backup_path(path: Path, *, label: str) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return path.with_name(f"{path.stem}_backup_{label}_{stamp}{path.suffix}")


def replace_with_fresh_workbook(
    workbook_path: Path,
    sheets: dict[str, list[list[object]]],
    *,
    backup_label: str,
    basic_formatting: bool = True,
    excel_tables: bool = True,
) -> tuple[Path, Path, dict[str, int]]:
    """
    Write a fresh workbook package, validate it, backup old workbook, replace.

    Returns (backup_path, new_temp_path, stats). The returned new_temp_path no
    longer exists after successful replacement; it is returned for logging/tests.
    """
    new_path = workbook_path.with_name(workbook_path.stem + ".new" + workbook_path.suffix)
    stats = write_fresh_workbook(
        new_path,
        sheets,
        basic_formatting=basic_formatting,
        excel_tables=excel_tables,
    )

    problems = validate_xlsx_basic(new_path, expected_sheets=set(sheets.keys()))
    if problems:
        raise RuntimeError("Fresh workbook validation failed:\n" + "\n".join(f"- {p}" for p in problems))

    backup_path = timestamped_backup_path(workbook_path, label=backup_label)
    shutil.move(str(workbook_path), str(backup_path))
    new_path.replace(workbook_path)

    return backup_path, new_path, stats

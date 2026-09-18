from openpyxl import Workbook

from finance_parser.budgeting.seed_account_balance import run


UNIFIED_COLUMNS = [
    "UnifiedID", "RawID", "SourceAccount", "SourceBank", "TransactionType",
    "Date", "Month", "Year", "Amount", "RawReceiver", "NormalizedReceiver",
    "Description", "Message", "Include", "Owner", "Supercategory",
    "Category", "Subcategory", "ReviewStatus", "Review/Notes",
]


def _write_unified(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "UnifiedTransactions"
    ws.append(UNIFIED_COLUMNS)
    for row in rows:
        ws.append([row.get(h, "") for h in UNIFIED_COLUMNS])
    wb.save(path)


def _feed(monkeypatch, answers):
    queue = list(answers)

    def fake_input(prompt=""):
        if not queue:
            raise AssertionError(f"Ran out of scripted answers at prompt: {prompt!r}")
        return queue.pop(0)

    monkeypatch.setattr("builtins.input", fake_input)


def test_run_with_no_qualifying_accounts_reports_and_exits(tmp_path):
    workbook = tmp_path / "ParsedTransactions.xlsx"
    settings_path = tmp_path / "settings.yaml"
    _write_unified(workbook, [
        {"SourceAccount": "PERSONAL", "SourceBank": "NORDEA", "Date": "2026-01-01", "Amount": -10},
    ])

    written = run(budgeting_workbook=workbook, settings_path=settings_path)

    assert written is False
    assert not settings_path.exists()


def test_run_end_to_end_writes_seed(tmp_path, monkeypatch):
    workbook = tmp_path / "ParsedTransactions.xlsx"
    settings_path = tmp_path / "settings.yaml"
    _write_unified(workbook, [
        {"SourceAccount": "HOUSEHOLD", "SourceBank": "OP", "Date": "2026-03-15", "Amount": -10},
    ])

    _feed(monkeypatch, [
        "1",       # select HOUSEHOLD (only account offered)
        "",        # accept default date
        "1234.56", # balance
        "y",       # confirm write
    ])

    written = run(budgeting_workbook=workbook, settings_path=settings_path)

    assert written is True
    content = settings_path.read_text(encoding="utf-8")
    assert "HOUSEHOLD" in content
    assert "1234.56" in content

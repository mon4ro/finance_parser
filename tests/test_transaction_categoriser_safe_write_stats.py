from finance_parser.budgeting.transaction_categoriser import (
    CategoriseStats,
    write_row_changes,
)


def test_categoriser_stats_fields_exist():
    stats = CategoriseStats()

    assert stats.rows_scanned == 0
    assert stats.rows_protected_by_comments == 0
    assert stats.rows_eligible == 0
    assert stats.rows_with_cleared_fields == 0
    assert stats.fields_cleared == 0
    assert stats.rows_changed_by_rules == 0
    assert stats.fields_changed_by_rules == 0
    assert stats.rows_changed_total == 0


class DummyCell:
    def __init__(self):
        self.value = None


class DummySheet:
    def __init__(self):
        self.cell_value = DummyCell()

    def cell(self, row, column):
        return self.cell_value


def test_write_row_changes_writes_cleaned_values():
    ws = DummySheet()
    write_row_changes(
        ws,
        2,
        {"Category": 1},
        {"Category": ("", "Restaurants")},
    )

    assert ws.cell_value.value == "Restaurants"

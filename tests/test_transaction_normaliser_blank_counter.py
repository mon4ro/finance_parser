import pandas as pd

from finance_parser.budgeting.transaction_normaliser import count_blank_normalized_receivers


def test_count_blank_normalized_receivers_counts_excel_style_blanks():
    df = pd.DataFrame({
        "NormalizedReceiver": ["PRISMA", "", None, "  ", "LIDL"],
    })

    assert count_blank_normalized_receivers(df) == 3


def test_count_blank_normalized_receivers_missing_column_is_zero():
    df = pd.DataFrame({"RawReceiver": ["A", "B"]})

    assert count_blank_normalized_receivers(df) == 0

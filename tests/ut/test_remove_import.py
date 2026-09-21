from finance_parser.utilities.remove_import import (
    PROFILES,
    apply_plan,
    build_plan,
)


def _budgeting_sheets():
    raw_headers = ["RawID", "SourceAccount", "SourceBank", "SourceFile", "BookingDate", "Amount"]
    raw_rows = [
        raw_headers,
        ["NOD-1", "CHILD", "NORDEA", "bad_import.csv", "2026-05-12", -4.5],
        ["NOD-2", "CHILD", "NORDEA", "bad_import.csv", "2026-06-01", 100],
        ["NOD-3", "CHILD", "NORDEA", "other_import.csv", "2026-06-02", -1],
    ]

    unified_headers = ["UnifiedID", "RawID", "SourceFile", "Date", "Amount", "Category", "Review/Notes"]
    unified_rows = [
        unified_headers,
        ["U-NOD-1", "NOD-1", "bad_import.csv", "2026-05-12", -4.5, "Restaurants", ""],
        ["U-NOD-2", "NOD-2", "bad_import.csv", "2026-06-01", 100, "", ""],
        ["U-NOD-3", "NOD-3", "other_import.csv", "2026-06-02", -1, "", ""],
    ]

    nordea_export_headers = ["SourceFile", "Kirjauspäivä", "Määrä"]
    nordea_export_rows = [
        nordea_export_headers,
        ["bad_import.csv", "2026/05/12", "-4,50"],
        ["bad_import.csv", "2026/06/01", "100"],
        ["other_import.csv", "2026/06/02", "-1"],
    ]

    import_log_headers = ["ImportRunID", "SourceFile", "SourceBank", "RowsNew"]
    import_log_rows = [
        import_log_headers,
        ["RUN-1", "bad_import.csv", "NORDEA", 2],
    ]

    return {
        "RawTransactions": raw_rows,
        "UnifiedTransactions": unified_rows,
        "NordeaRawExport": nordea_export_rows,
        "ImportLog": import_log_rows,
    }


def test_build_plan_finds_matching_rows_across_all_sheets():
    sheets = _budgeting_sheets()
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    assert len(plan.raw_matched_idx) == 2
    assert len(plan.classified_matched_idx) == 2
    assert plan.raw_export_sheet == "NordeaRawExport"
    assert len(plan.raw_export_matched_idx) == 2
    assert not plan.blocking_reviewed
    assert not plan.blocking_split


def test_build_plan_does_not_touch_other_source_files():
    sheets = _budgeting_sheets()
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    matched_raw_ids = {plan.raw_records[i]["RawID"] for i in plan.raw_matched_idx}
    assert matched_raw_ids == {"NOD-1", "NOD-2"}
    assert "NOD-3" not in matched_raw_ids


def test_build_plan_blocks_on_non_blank_review_notes():
    sheets = _budgeting_sheets()
    sheets["UnifiedTransactions"][1][-1] = "checked this one manually"  # Review/Notes on U-NOD-1
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    assert len(plan.blocking_reviewed) == 1
    assert plan.blocking_reviewed[0]["UnifiedID"] == "U-NOD-1"


def test_build_plan_blocks_on_split_child_and_parent():
    sheets = _budgeting_sheets()
    # Make U-NOD-1 look like an already-split child row (still matched via
    # its own SourceFile, unrelated to the UnifiedID rename).
    sheets["UnifiedTransactions"][1][0] = "U-NOD-1-S01"
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    assert len(plan.blocking_split) == 1


def test_build_plan_blocks_when_row_is_a_split_parent_with_children():
    sheets = _budgeting_sheets()
    # Add a child row (different SourceFile, so it wouldn't otherwise be touched)
    # whose UnifiedID marks U-NOD-2 as a split parent.
    sheets["UnifiedTransactions"].append(
        ["U-NOD-2-S01", "NOD-99", "somewhere_else.csv", "2026-06-01", 50, "", ""]
    )
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    blocked_uids = {r["UnifiedID"] for r in plan.blocking_split}
    assert "U-NOD-2" in blocked_uids


def test_apply_plan_removes_only_matched_rows_and_leaves_import_log_untouched():
    sheets = _budgeting_sheets()
    profile = PROFILES["budgeting"]()
    plan = build_plan(sheets, profile, "bad_import.csv")

    stats = apply_plan(sheets, plan)

    assert stats["RawTransactions"] == (3, 1)
    assert stats["UnifiedTransactions"] == (3, 1)
    assert stats["NordeaRawExport"] == (3, 1)
    assert "ImportLog" not in stats
    # ImportLog sheet itself is never rewritten by apply_plan.
    assert sheets["ImportLog"] == _budgeting_sheets()["ImportLog"]

    # The surviving row is the other_import.csv one, in every touched sheet.
    remaining_raw_source_files = [row[3] for row in sheets["RawTransactions"][1:]]
    assert remaining_raw_source_files == ["other_import.csv"]


def test_investment_profile_uses_comments_as_review_gate_and_has_no_split_check():
    raw_headers = ["InvestmentRawID", "Broker", "SourceFile", "TradeDate", "CashAmount"]
    raw_rows = [
        raw_headers,
        ["INV-1", "EVLI", "bad_export.csv", "2026-05-12", -100],
    ]
    tx_headers = ["InvestmentTransactionID", "InvestmentRawID", "SourceFile", "TradeDate", "CashAmount", "Comments"]
    tx_rows = [
        tx_headers,
        ["T-1", "INV-1", "bad_export.csv", "2026-05-12", -100, "manually annotated"],
    ]
    sheets = {
        "RawInvestmentTransactions": raw_rows,
        "InvestmentTransactions": tx_rows,
        "EvliRawExport": [["SourceFile"], ["bad_export.csv"]],
        "InvestmentImportLog": [["SourceFile"], ["bad_export.csv"]],
    }
    profile = PROFILES["investments"]()

    plan = build_plan(sheets, profile, "bad_export.csv")

    assert len(plan.classified_matched_idx) == 1
    assert len(plan.blocking_reviewed) == 1
    assert not plan.blocking_split  # investments profile never populates this


def test_build_plan_leaves_raw_export_sheet_alone_when_source_bank_is_ambiguous():
    sheets = _budgeting_sheets()
    # Make the two matched raw rows disagree on SourceBank.
    sheets["RawTransactions"][2][2] = "OP"
    profile = PROFILES["budgeting"]()

    plan = build_plan(sheets, profile, "bad_import.csv")

    assert plan.raw_export_sheet is None
    assert plan.raw_export_matched_idx == []

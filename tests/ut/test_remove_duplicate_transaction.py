from finance_parser.budgeting.remove_duplicate_transaction import (
    apply_plan,
    build_plan,
    verify_same_transaction,
)


def _sheets():
    raw_headers = [
        "RawID", "SourceAccount", "SourceBank", "SourceFile", "ImportedAt",
        "BookingDate", "Amount", "RawReceiver", "Message", "ArchiveID",
    ]
    raw_rows = [
        raw_headers,
        ["OP-old", "CHILD", "OP", "file_v1.csv", "2026-01-15 10:00:00",
         "2026-03-05", 250.0, "SENDER NAME", "A gift", "ARCHIVE-1"],
        ["OP-new", "CHILD", "OP", "file_v1.csv", "2026-02-01 09:00:00",
         "2026-03-05", 250.0, "SENDER NAME", "A gift", "ARCHIVE-1"],
        ["OP-unrelated", "CHILD", "OP", "file_v1.csv", "2026-02-01 09:00:00",
         "2026-08-12", -20.0, "SOME SHOP", "", "ARCHIVE-2"],
    ]

    unified_headers = ["UnifiedID", "RawID", "Amount", "Supercategory", "Category", "Subcategory", "Review/Notes"]
    unified_rows = [
        unified_headers,
        ["U-OP-old", "OP-old", 250.0, "", "", "", ""],
        ["U-OP-new", "OP-new", 250.0, "", "", "", ""],
        ["U-OP-unrelated", "OP-unrelated", -20.0, "EXPENSES", "Shopping", "Clothes", ""],
    ]

    return {"RawTransactions": raw_rows, "UnifiedTransactions": unified_rows}


def test_verify_same_transaction_passes_when_only_identity_columns_differ():
    records = [
        {"RawID": "A", "ImportedAt": "t1", "Amount": 250.0, "RawReceiver": "X"},
        {"RawID": "B", "ImportedAt": "t2", "Amount": 250.0, "RawReceiver": "X"},
    ]
    assert verify_same_transaction(records) == []


def test_verify_same_transaction_flags_real_differences():
    records = [
        {"RawID": "A", "ImportedAt": "t1", "Amount": 250.0, "RawReceiver": "X"},
        {"RawID": "B", "ImportedAt": "t2", "Amount": 250.0, "RawReceiver": "Y"},
    ]
    assert verify_same_transaction(records) == ["RawReceiver"]


def test_build_plan_verifies_and_keeps_the_oldest():
    plan = build_plan(_sheets(), ["OP-old", "OP-new"])

    assert plan.mismatches == []
    assert plan.keep_raw_id == "OP-old"
    assert plan.remove_raw_ids == ["OP-new"]
    assert not plan.blocking_reviewed


def test_build_plan_refuses_a_real_mismatch():
    plan = build_plan(_sheets(), ["OP-old", "OP-unrelated"])

    assert plan.mismatches
    assert plan.keep_raw_id is None


def test_build_plan_reports_missing_raw_id():
    plan = build_plan(_sheets(), ["OP-old", "OP-does-not-exist"])

    assert plan.not_found == ["OP-does-not-exist"]
    assert plan.keep_raw_id is None


def test_build_plan_blocks_on_non_blank_review_notes():
    sheets = _sheets()
    sheets["UnifiedTransactions"][2][-1] = "checked this by hand"  # Review/Notes on U-OP-new
    plan = build_plan(sheets, ["OP-old", "OP-new"])

    assert len(plan.blocking_reviewed) == 1
    assert plan.blocking_reviewed[0]["UnifiedID"] == "U-OP-new"


def test_build_plan_prefers_the_row_that_matches_current_code_over_the_oldest():
    """
    Real incident this guards against: a parsing-logic fix landed between
    two imports of the same real transaction, so the OLDER row was hashed
    under the old, now-dead logic and the NEWER row happens to match what
    current code recomputes. Plain "keep oldest" would keep the wrong one
    here - the one that will never be recomputed again - so the very next
    re-import of the source file would recreate this exact duplicate.
    """
    raw_headers = [
        "RawID", "SourceAccount", "SourceBank", "SourceFile", "ImportedAt",
        "ValueDate", "Amount", "RawReceiver", "Description", "Reference", "Message", "ArchiveID",
    ]
    # This is the real hash op.make_raw_id() produces for these exact field
    # values today - computed once via the real function, not invented.
    current_code_raw_id = "OP-07488fe2bd0d"
    raw_rows = [
        raw_headers,
        ["OP-old-wrong-hash", "CHILD", "OP", "file.csv", "2026-07-01 00:00:00",
         "2026-03-05", 250.0, "SENDER NAME", "", "", "A gift", "ARCHIVE-1"],
        [current_code_raw_id, "CHILD", "OP", "file.csv", "2026-09-01 00:00:00",
         "2026-03-05", 250.0, "SENDER NAME", "", "", "A gift", "ARCHIVE-1"],
    ]
    unified_headers = ["UnifiedID", "RawID", "Amount", "Review/Notes"]
    unified_rows = [
        unified_headers,
        ["U-OP-old-wrong-hash", "OP-old-wrong-hash", 250.0, ""],
        [f"U-{current_code_raw_id}", current_code_raw_id, 250.0, ""],
    ]
    sheets = {"RawTransactions": raw_rows, "UnifiedTransactions": unified_rows}

    plan = build_plan(sheets, ["OP-old-wrong-hash", current_code_raw_id])

    assert plan.mismatches == []
    assert plan.keep_raw_id == current_code_raw_id
    assert plan.remove_raw_ids == ["OP-old-wrong-hash"]
    assert "matches what current code recomputes" in plan.keep_reason


def test_apply_plan_removes_only_the_confirmed_duplicate():
    sheets = _sheets()
    plan = build_plan(sheets, ["OP-old", "OP-new"])

    stats = apply_plan(sheets, plan)

    assert stats["RawTransactions"] == (3, 2)
    assert stats["UnifiedTransactions"] == (3, 2)
    remaining_raw_ids = {row[0] for row in sheets["RawTransactions"][1:]}
    assert remaining_raw_ids == {"OP-old", "OP-unrelated"}
    remaining_unified_raw_ids = {row[1] for row in sheets["UnifiedTransactions"][1:]}
    assert remaining_unified_raw_ids == {"OP-old", "OP-unrelated"}

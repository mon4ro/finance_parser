import pandas as pd

from finance_parser.common import drop_duplicates_against_existing_only


def _combined(existing_keys, new_keys):
    """Build a combined frame matching the real call sites' own shape:
    pd.concat([existing, new_rows], ignore_index=True) then a
    _CanonicalTransactionKey column added - the function only reads the key
    column and the existing/new row-count split, so a bare key list is
    enough to exercise it directly."""
    keys = list(existing_keys) + list(new_keys)
    return pd.DataFrame({
        "Label": [f"row{i}" for i in range(len(keys))],
        "_CanonicalTransactionKey": keys,
    }, index=range(len(keys))), len(existing_keys)


def test_two_new_rows_sharing_a_key_are_both_kept():
    """
    Real motivating case: two genuinely separate real transactions (two
    airline seats billed as two separate identical charges the same day;
    two identically-priced drinks at a bar, back to back) share the exact
    same date/amount/receiver in a single fresh import from a bank with no
    richer per-row reference (NORWEGIAN). Canonical-key equality between two
    rows that both just came from this run's own import is not proof either
    is a duplicate.
    """
    combined, n_existing = _combined(existing_keys=[], new_keys=["K1", "K1", "K1"])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 3


def test_new_row_matching_something_already_existing_is_dropped():
    """The one case this function must still catch: a fresh import re-adding
    a transaction that was already imported in an earlier run (e.g. an
    overlapping-date-range re-export, or literally re-importing the same
    file)."""
    combined, n_existing = _combined(existing_keys=["K1"], new_keys=["K1"])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 1
    assert result.iloc[0]["Label"] == "row0"  # the existing row, not the new one


def test_new_rows_exceeding_the_existing_count_for_the_same_key_are_all_kept():
    """
    A subtler real case: yesterday's import correctly kept 3 real bar
    purchases sharing a key. Today's re-export of the same file re-parses
    those same 3 - they must all still collapse down to the 3 already in
    "existing" (not balloon to 6), while a 4th, genuinely new transaction
    sharing that key stays a known, accepted limitation (set membership
    can't distinguish "still exactly 3" from "a real 4th happened") rather
    than a regression this function needs to solve.
    """
    combined, n_existing = _combined(existing_keys=["K1", "K1", "K1"], new_keys=["K1", "K1", "K1"])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 3
    assert all(result["Label"] == [f"row{i}" for i in range(3)])


def test_existing_rows_sharing_a_key_are_never_re_collapsed():
    """
    An earlier version of this function still deduplicated existing rows
    among themselves as a leftover safety net - but that's the exact same
    flawed reasoning this function exists to fix, just applied a run later:
    once several genuinely separate real transactions sharing a key have
    been correctly accepted (by an earlier run), a later run must not
    silently shrink them back down just because they still share that key.
    """
    combined, n_existing = _combined(existing_keys=["K1", "K1", "K1"], new_keys=[])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 3


def test_distinct_keys_are_all_kept_regardless_of_origin():
    combined, n_existing = _combined(existing_keys=["A"], new_keys=["B", "C"])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 3


def test_empty_combined_frame_returns_empty():
    combined, n_existing = _combined(existing_keys=[], new_keys=[])
    result = drop_duplicates_against_existing_only(combined, n_existing)
    assert len(result) == 0
